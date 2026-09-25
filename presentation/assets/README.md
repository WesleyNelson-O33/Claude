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
- `cts-logo-20yrs-white.svg`: the 20-year anniversary lock-up currently on
  ctsav.com.au (`/wp-content/uploads/cts-assets/CTS_Logo_20Anniv_RGB_White_FullColour.svg`),
  white wordmark with the "20 YRS" mark in mint, lavender and rose. Not used
  on the layouts; swap it in on the covers if you want the anniversary version.

The build script places `cts-logo.png` on white layouts and `cts-logo-dark.png`
on the black and navy ones. Run `node presentation/build_template.js` after
changing any of them.
