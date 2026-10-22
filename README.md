# CD Collection

Searchable personal catalogue of CDs and SACDs: shelf location for each, plus label, catalogue number, barcode and pressing for each copy.

Live site: https://comstocd.github.io/cd-collection/

- `index.html` — the page (List and Shelf views; built for iPad Air on iOS 12 Safari)
- `collection.js` — the data (`window.COLLECTION`): `items` (one entry per title, with a `copies` list) and `cubbies` (the shelf grid)
- `sw.js` — offline cache and update detection; a new `BUILD` number triggers "Update available"
- `tools/build.py` — regenerates `collection.js` from `CD_Catalogue.xlsx` (sheets: Shelf plan, Picard Jan 2023, Photo batch 1-6)
- `tools/deploy.sh` — rebuild, bump BUILD, commit and push


