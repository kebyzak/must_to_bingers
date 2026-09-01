# Must → Bingers Converter

Converts your Must watch history into a Trakt-compatible export that can be imported into Bingers.

## 1. Export your Must data

Go to:

https://mustapp-exporter.pages.dev/

Follow the instructions on the website and export your Must data as an Excel (`.xlsx`) file.

## 2. Rename the Excel file

Rename the downloaded file to:

`mustapp.xlsx`

Place it in the same folder as `main.js`.


## 3. Install dependencies

Open a terminal in the project folder and run:

npm install

If the dependencies are not already listed in `package.json`, run:

npm install xlsx axios adm-zip

## 4. Add your TMDB API key

Open `main.js` and find:

Replace it with your TMDB API Key (v3):

const TMDB_API_KEY = "YOUR_TMDB_API_KEY";

## 5. Run the converter

Run:

node main.js

The script will:

- Read movies and shows from `mustapp.xlsx`
- Match them with TMDB
- Reconstruct watched TV episodes
- Generate a Trakt-compatible export
- Create `trakt-export.zip`

## 6. Import into Bingers

Open Bingers and go to the import page.

Select the Trakt import option and upload:

`trakt-export.zip`

Bingers should detect the export and import your watch history.

## Important

Must only stores the total number of watched episodes for a show in its export. It does not provide the exact individual episodes watched.

Because of this, the converter assumes that the first N episodes were watched in chronological season/episode order.

For example, if Must says that 18 episodes of a show were watched, the converter marks the first 18 regular episodes as watched.

Season 0 / Specials are excluded.