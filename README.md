# trim_ical

**A single PHP file that trims bloated iCal feeds down to a usable window** so slow calendar clients (Home Assistant, e-paper displays, low-RAM widgets, anything that polls and parses in-process) can actually keep up.

Drop it on any PHP host. Point your client at it instead of directly at Outlook / Exchange / Google Calendar. Get back a small, well-formed feed with recurring events, cancellations, timezones, and DST handled correctly.

```
+--------------------+      +------------------+      +---------------------+
|  Outlook / GCal    |  ->  |  trim_ical.php   |  ->  |  Home Assistant     |
|  (12 MB, 2018+)    |      |   (~150 KB out)  |      |  (parses in ~ms)    |
+--------------------+      +------------------+      +---------------------+
```

---

## Why this exists

Home Assistant's iCal implementation is garbage and always parses the whole calendar in a way that not only overloads the backend with processing but even kills your front-end performance if you have too many events in your massive calendar.

This script trims down the entire ical file to just current events on-the-fly so it can be used with Home Assistant without killing your Raspberry pi.

---

## Quick start

You need PHP 7.4+ with the cURL extension. Most shared hosts qualify.

**1. Grab the file.**

```bash
curl -O https://raw.githubusercontent.com/permissionBRICK/trim_ical/main/trim_ical.php
```

Upload it to your webroot however you normally do.

**2. Generate a token.**

```bash
openssl rand -hex 32
```

**3. Edit `trim_ical.php`** and replace the placeholder token near the top:

```php
$expectedToken = 'paste-your-hex-token-here';
```

The script will refuse to run until you do this.

**4. Test it.**

```bash
export HOST='example.com'                                       # the host where you uploaded trim_ical.php
export URL='https://outlook.office365.com/owa/calendar/.../calendar.ics'
export token='your-hex-token'
export tz='America/Los_Angeles'                                 # optional, IANA zone for floating times

URL_ENC=$(python3 -c 'import os,urllib.parse; print(urllib.parse.quote(os.environ["URL"], safe=""))')
curl -s "https://${HOST}/trim_ical.php?url=${URL_ENC}&token=${token}&tz=${tz}" | head -5
```

You should get back something starting with `BEGIN:VCALENDAR`. That's it. 

Now take that final encoded url and paste it into a Home Assistant Remote Calendar integration and enjoy a lag-free experience.

---

## Query parameters

| Parameter | Required | Notes |
|---|---|---|
| `url` | yes | The upstream `.ics` URL, **percent-encoded**. |
| `token` | yes | Must match `$expectedToken`. |
| `tz` | no | IANA timezone (e.g. `Europe/Berlin`) used for floating times when the source feed has no `X-WR-TIMEZONE`. |

## Contributing

This was a one-shot vibecode project. If you're looking for an actively maintained version, go [here](https://github.com/ChrisCarini/trim_ical)

## License

[MIT](LICENSE)
