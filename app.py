"""Flask backend for the Events at Equestrian Woods landing page.

Serves templates/index.html at / (the old /preopenreservations link redirects there)
and receives its inquiry form at POST /api/inquiry,
emailing each inquiry to the venue team using templates/<inquiry_type>_email.html.
"""

import json
import logging
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from flask import Flask, jsonify, redirect, render_template, request, url_for
from flask_cors import CORS

load_dotenv()
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024

cors_origins = os.environ.get("CORS_ORIGINS", "*")
CORS(app, resources={r"/api/*": {
    "origins": cors_origins if cors_origins == "*" else [o.strip() for o in cors_origins.split(",")],
}})

# Keep in sync with MIN_DATE in templates/index.html.
MIN_DATE = date(2026, 11, 1)
VENUE_TZ = ZoneInfo("America/New_York")  # Nicholasville, KY
EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
# Public address of the site; keep in sync with the canonical link in templates/index.html.
CANONICAL_HOST = "preopenreservations.eventsatequestrian.com"


def clean(data, key, limit=200, multiline=False):
    """Return a trimmed string field from the JSON body; anything else becomes ""."""
    value = data.get(key)
    if not isinstance(value, str):
        return ""
    if not multiline:
        value = " ".join(value.split())
    return value.strip()[:limit]


def format_date(d):
    return f"{d:%A, %B} {d.day}, {d.year}"


def bad_request(error):
    return jsonify(ok=False, error=error), 400


def post(url, data, headers):
    """POST and return the response body; on an HTTP error, include Microsoft's reason in the exception."""
    try:
        with urllib.request.urlopen(urllib.request.Request(url, data=data, headers=headers), timeout=20) as resp:
            return resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"{url} returned {e.code}: {e.read().decode(errors='replace')}") from e


def send_email(subject, html, reply_to):
    """Send to MAIL_TO from the MAIL_FROM mailbox via Microsoft Graph, signed in as the Entra app."""
    token = json.loads(post(
        f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}/oauth2/v2.0/token",
        urllib.parse.urlencode({
            "grant_type": "client_credentials",
            "client_id": os.environ["AZURE_CLIENT_ID"],
            "client_secret": os.environ["AZURE_CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
        }).encode(),
        {"Content-Type": "application/x-www-form-urlencoded"},
    ))["access_token"]

    recipients = [a.strip() for a in os.environ["MAIL_TO"].split(",") if a.strip()]
    post(
        f"https://graph.microsoft.com/v1.0/users/{urllib.parse.quote(os.environ['MAIL_FROM'])}/sendMail",
        json.dumps({"message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": a}} for a in recipients],
            "replyTo": [{"emailAddress": {"address": reply_to}}],
        }}).encode(),
        {"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )


@app.before_request
def redirect_to_canonical_host():
    # Send the onrender.com address to the custom domain so Google indexes only one copy.
    # /healthz is exempt so Render's health checks aren't redirected.
    if request.host.endswith(".onrender.com") and request.path != "/healthz":
        return redirect(f"https://{CANONICAL_HOST}{request.full_path.rstrip('?')}", code=301)


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/robots.txt")
def robots():
    body = f"User-agent: *\nAllow: /\n\nSitemap: https://{CANONICAL_HOST}/sitemap.xml\n"
    return body, {"Content-Type": "text/plain; charset=utf-8"}


@app.get("/sitemap.xml")
def sitemap():
    body = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        f"  <url><loc>https://{CANONICAL_HOST}/</loc></url>\n"
        "</urlset>\n"
    )
    return body, {"Content-Type": "application/xml; charset=utf-8"}


@app.get("/preopenreservations", strict_slashes=False)
def preopenreservations():
    # Former address of the page; keep links already shared working.
    return redirect(url_for("index"))


@app.get("/healthz")
def healthz():
    return jsonify(ok=True)


@app.post("/api/inquiry")
def inquiry():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return bad_request("Expected a JSON object.")

    # Honeypot: the page hides this field from people, so only bots fill it in.
    # Report success so they don't retry.
    if clean(data, "company_website"):
        return jsonify(ok=True)

    inquiry_type = clean(data, "inquiry_type")
    if inquiry_type not in ("reservation", "general"):
        return bad_request("Unknown inquiry type.")

    now = datetime.now(VENUE_TZ)
    fields = {
        "submitted_at": f"{format_date(now)} at {now.hour % 12 or 12}:{now:%M %p %Z}",
        "name": clean(data, "name"),
        "email": clean(data, "email"),
        "phone": clean(data, "phone"),
        "message": clean(data, "message", limit=10_000, multiline=True),
    }
    if not fields["name"]:
        return bad_request("Name is required.")
    if not EMAIL_RE.match(fields["email"]):
        return bad_request("A valid email address is required.")

    if inquiry_type == "general":
        if not fields["message"]:
            return bad_request("Please include a message.")
        subject = f"General inquiry: {fields['name']}"
    else:
        fields["event_type"] = clean(data, "event_type")
        fields["guest_count"] = clean(data, "guest_count")
        if not fields["event_type"] or not fields["guest_count"]:
            return bad_request("Event type and guest count are required.")

        flexible = clean(data, "dates_flexible") == "yes"
        for key in ("preferred_date", "backup_date"):
            raw = clean(data, key)
            try:
                day = date.fromisoformat(raw) if raw else None
            except ValueError:
                return bad_request("Dates must be in YYYY-MM-DD format.")
            if day and day < MIN_DATE:
                return bad_request(f"Dates must be on or after {format_date(MIN_DATE)}.")
            fields[key] = format_date(day) if day else ""
        if not fields["preferred_date"] and not flexible:
            return bad_request("A preferred date is required unless your dates are flexible.")

        fields["dates_flexible"] = "Yes" if flexible else "No"
        fields["tour_requested"] = "Yes" if clean(data, "tour_requested") == "yes" else "No"
        fields["heard_about"] = clean(data, "heard_about")
        subject = f"Reservation inquiry: {fields['name']} – {fields['event_type']}"

    shown = {key: value or "—" for key, value in fields.items()}
    html = render_template(f"{inquiry_type}_email.html", **shown)
    text = "\n".join(
        f"{key.replace('_', ' ').capitalize()}: {value}"
        for key, value in shown.items() if key != "message"
    ) + f"\n\nMessage:\n{shown['message']}"

    try:
        send_email(subject, html, reply_to=fields["email"])
    except Exception:
        # Log the contents so the inquiry isn't lost if email is misconfigured or down.
        app.logger.exception("Could not email %s inquiry:\n%s", inquiry_type, text)
        return jsonify(ok=False, error="We couldn't send your inquiry. Please try again later."), 502

    app.logger.info("Sent %s inquiry from %s", inquiry_type, fields["email"])
    return jsonify(ok=True)


if __name__ == "__main__":
    app.run(debug=True)
