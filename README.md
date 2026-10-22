# CD Collection

Searchable catalogue of David Comstock's CDs and SACDs: which cubby of the 4 x 6 shelf each one lives in, plus label, catalogue number, barcode and pressing for each copy.

Live site: https://comstocd.github.io/cd-collection/

- `index.html` — the page (List and Shelf views; built for iPad Air on iOS 12 Safari)
- `collection.js` — the data (`window.COLLECTION`): `items` (one entry per title, with a `copies` list) and `cubbies` (the shelf grid)
- `sw.js` — offline cache and update detection; a new `BUILD` number triggers "Update available"
- `tools/build.py` — regenerates `collection.js` from `CD_Catalogue.xlsx` (sheets: Shelf plan, Picard Jan 2023, Photo batch 1-6)
- `tools/deploy.sh` — rebuild, bump BUILD, commit and push

Pressing details come from the MusicBrainz release matched when each disc was ripped (January 2023) or from spine photos. Entries flagged `check` have partly unreadable spines.
