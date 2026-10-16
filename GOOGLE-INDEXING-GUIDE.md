# Getting OKV SACCO Subscribe Indexed on Google

Two files are already in this project's root for you: **`robots.txt`** and
**`sitemap.xml`**, both pointing at `https://www.okvtechnology.com`. `index.html`
also now has a meta description, canonical tag, and Open Graph/Twitter card
tags — without these, Google writes its own (often awkward) snippet for your
search result, and links shared on WhatsApp/Facebook/Twitter show no preview
card at all.

**If your live domain is different from `www.okvtechnology.com`**, find-and-replace
that domain in `robots.txt`, `sitemap.xml`, and `index.html`'s `<head>` before
publishing.

## Steps

1. **Confirm the files are live.** After deploying, visit
   `https://www.okvtechnology.com/robots.txt` and
   `https://www.okvtechnology.com/sitemap.xml` in a browser — both should load
   as plain text/XML, not a 404.
2. **Set up Google Search Console.** Go to https://search.google.com/search-console,
   add your site as a property. For a static site like this, use the **HTML tag**
   verification method: Search Console gives you a
   `<meta name="google-site-verification" content="...">` tag — paste it into
   `index.html`'s `<head>`, deploy, then click Verify.
3. **Submit the sitemap.** In Search Console, go to Sitemaps (left sidebar),
   enter `sitemap.xml`, click Submit.
4. **Request indexing for your most important pages.** Use the URL Inspection
   tool (top search bar in Search Console), paste the homepage URL, click
   "Request Indexing." Repeat for `pricing.html` and `demo.html` — those matter
   most for conversions.
5. **Wait.** Initial indexing typically takes a few days to a couple of weeks.
   The Coverage report in Search Console shows what's been indexed and flags
   any errors.

## What's excluded from indexing, and why

`robots.txt` disallows `login.html`, `reset-password.html`, `dashboard.html`,
and `super-admin.html` — these are account/app pages with no unique public
content, and indexing them would only dilute crawl budget or produce
confusing search results (someone landing on a bare login form isn't useful
to them or to your ranking). `signup.html`, `demo.html`, `pricing.html`, and
the three legal pages remain crawlable and are listed in the sitemap.

## Keeping this current

If you add new public marketing pages later, add a `<url>` entry for each to
`sitemap.xml` and re-submit it in Search Console — Google won't discover a
new page automatically just because it's linked from the homepage, though
that does help too.
