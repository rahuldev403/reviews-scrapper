# Google Play 1-Star Review Scraper

Scrapes **all** reviews of a given star rating from a Google Play app and exports
them to a formatted Excel sheet.

## How it works

Google Play loads reviews dynamically from an internal API rather than putting them
in the page HTML. This tool uses the [`google-play-scraper`](https://pypi.org/project/google-play-scraper/)
library, which calls that same API. It:

1. Requests reviews filtered **server-side** to one star rating (`filter_score_with`).
2. Walks every page using the API's **continuation token** until none are left.
3. De-duplicates and writes the result to `.xlsx` (frozen header, autofilter,
   wrapped text, sensible column widths).

## Setup

```bash
pip install -r requirements.txt
```

## Usage

Default — all 1-star reviews of `org.grow90.whatsub` (India / English):

```bash
python scrape_reviews.py
```

Options:

```bash
python scrape_reviews.py \
  --app-id org.grow90.whatsub \   # the ?id= value in the Play Store URL
  --score 1 \                     # star rating to scrape (1-5)
  --country in \                  # store country
  --lang en \                     # review language
  --output my_reviews.xlsx        # optional output path
```

If `--output` is omitted, the file is named
`reviews_<appid>_<score>star_<timestamp>.xlsx`.

## Output columns

| Column | Description |
|--------|-------------|
| Review ID | Unique id for the review |
| User Name | Reviewer's display name |
| Rating | Star rating (1–5) |
| Review | Review text |
| Thumbs Up | How many users found it helpful |
| App Version | App version the review was written on |
| Date | When the review was posted |
| Developer Reply | The developer's response, if any |
| Reply Date | When the developer replied |

## Notes

- The Play Store API returns the reviews it chooses to surface for the given
  country/language; running with different `--country` / `--lang` values can
  return additional reviews from other locales.
- A short delay is added between pages to be polite to the endpoint.
# reviews-scrapper
