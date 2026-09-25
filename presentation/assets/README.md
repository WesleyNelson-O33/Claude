# CTS logo files

All files are the official CTS Seamless AV logo, taken from the company's own
website. Sources:

- `cts-logo.png` (charcoal on transparent, 1377 x 936) and `cts-logo-dark.png`
  (white on transparent, 1377 x 936): the "CTS Logos_Seamless AV_Charcoal_RGB"
  and "_White_RGB" exports served by ctsav.com.au in 2022 to 2024
  (Webflow CDN, `cdn.prod.website-files.com/62e1e7bd826ffbfc00162294/...`).
- `cts-logo-white.svg`: the vector "CTS_Logo_White.svg" from the 2022 site.
  `cts-logo-black.svg` is the same vector with the fill changed to brand
  black 121820.
- `cts-wordmark.png` / `cts-wordmark-dark.png`: the "cts" letters alone, cropped
  from the exports above.
- `cts-production*.png`, `cts-consulting*.png`, `cts-support*.png`: division
  lock-ups built from the wordmark plus the division name in Tomato Grotesk
  Regular (the face used for "Seamless AV" on the master logo). No suffix =
  charcoal wordmark, dark division colour, for white backgrounds. `-dark` =
  white wordmark, light division colour, for black backgrounds. `-white` =
  all white, for the division's own colour.
- `cts-logo-20yrs-white.svg`: the 20-year anniversary lock-up currently on
  ctsav.com.au (`/wp-content/uploads/cts-assets/CTS_Logo_20Anniv_RGB_White_FullColour.svg`),
  white wordmark with the "20 YRS" mark in mint, lavender and rose. Rasterised
  as `cts-logo-20yrs-dark.png` (used on the 20 years cover) and
  `cts-logo-20yrs.png` (charcoal letters, for white backgrounds);
  `cts-logo-20yrs-black.svg` is the vector with black letters.

The build script picks the right file for each layout by name. Run `node presentation/build_template.js` after
changing any of them.
