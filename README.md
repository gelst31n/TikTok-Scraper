
# Tiktok Sound Scraper

## Scraper
TikTok Scraper using TikAPI that pulls all videos made with a provided Sound ID. Data scraped includes the following fields: username, secUid, link,nickname, followers, posts under sound, bio,posts in last 28 days, median views in last 28 days, mean views in last 28 days, median likes in last 28 days, median comments in last 28 days, median views per follower, tags, index

## Cleaner
Takes output from scraper.py. Filters data, calculates key metrics, extracts instagram handles or email from text in bio,  automatically applies hyperlinks in the username column and formats data as output in an xlsx file.
