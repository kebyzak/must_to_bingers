# Must → Bingers Converter

Converts your Must watch history, ratings, and Want list into exports that can be imported into Bingers.

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
- Transfer movie and show ratings through the Trakt export; Bingers converts them to its star scale
- Add entries from the `Want` sheet to the Bingers library as not started
- Add watched movies and shows to the Bingers library independently from their Trakt watch history
- Create `trakt-export.zip` for watched history and ratings
- Create `bingers-library.zip` for the Bingers library/bookmarks
- Remove the temporary output folders after the ZIP files are created

Import the ZIP files into Bingers separately, in the order shown below. The library archive uses the TV Time Out format to create the list/bookmark state; the Trakt archive preserves watched status and ratings.

## 6. Import into Bingers

Open Bingers and go to the import page.

Upload these archives one at a time:

`trakt-export.zip`

`bingers-library.zip`

Bingers detects the source format from each archive. The Trakt archive contains both watched history and ratings.

## Important

Must only stores the total number of watched episodes for a show in its export. It does not provide the exact individual episodes watched.

Because of this, the converter assumes that the first N episodes were watched in chronological season/episode order.

For example, if Must says that 18 episodes of a show were watched, the converter marks the first 18 regular episodes as watched.

Season 0 / Specials are excluded.
