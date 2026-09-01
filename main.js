const XLSX = require("xlsx");
const axios = require("axios");
const AdmZip = require("adm-zip");
const fs = require("fs");
const path = require("path");

// =====================================================
// CONFIG
// =====================================================

const TMDB_API_KEY = "eb4a12294b26b9623b2a9196ec8f5f0b";

const INPUT_FILE = "mustapp.xlsx";

const OUT_DIR = path.join(__dirname, "trakt-export");
const OUT_ZIP = path.join(__dirname, "trakt-export.zip");

const tmdb = axios.create({
  baseURL: "https://api.themoviedb.org/3",
  params: {
    api_key: TMDB_API_KEY,
  },
});

// =====================================================
// HELPERS
// =====================================================

function normalizeTitle(value) {
  return String(value || "").trim();
}

function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (!parsed) return null;

    return new Date(
      Date.UTC(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        parsed.H || 0,
        parsed.M || 0,
        parsed.S || 0
      )
    );
  }

  const d = new Date(value);

  if (Number.isNaN(d.getTime())) {
    return null;
  }

  return d;
}

function toISOString(value) {
  const date = parseDate(value);

  if (!date) {
    return new Date().toISOString();
  }

  return date.toISOString();
}

function getYear(value) {
  const date = parseDate(value);

  if (!date) {
    return undefined;
  }

  return date.getUTCFullYear();
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// =====================================================
// TMDB
// =====================================================

async function findMovie(title, year) {
  console.log(`Searching movie: ${title} (${year || "?"})`);

  const params = {
    query: title,
  };

  if (year) {
    params.year = year;
  }

  let response = await tmdb.get("/search/movie", {
    params,
  });

  let results = response.data.results || [];

  if (!results.length && year) {
    response = await tmdb.get("/search/movie", {
      params: {
        query: title,
      },
    });

    results = response.data.results || [];
  }

  if (!results.length) {
    throw new Error(`Movie not found in TMDB: ${title}`);
  }

  let movie = results[0];

  if (year) {
    const exact = results.find((item) => {
      if (!item.release_date) return false;

      return (
        Number(item.release_date.slice(0, 4)) ===
        Number(year)
      );
    });

    if (exact) {
      movie = exact;
    }
  }

  console.log(
    `  -> ${movie.title} (${movie.release_date || "?"}), TMDB=${movie.id}`
  );

  return movie;
}

async function findShow(title, year) {
  console.log(`Searching show: ${title} (${year || "?"})`);

  const params = {
    query: title,
  };

  if (year) {
    params.first_air_date_year = year;
  }

  let response = await tmdb.get("/search/tv", {
    params,
  });

  let results = response.data.results || [];

  if (!results.length && year) {
    response = await tmdb.get("/search/tv", {
      params: {
        query: title,
      },
    });

    results = response.data.results || [];
  }

  if (!results.length) {
    throw new Error(`Show not found in TMDB: ${title}`);
  }

  let show = results[0];

  if (year) {
    const exact = results.find((item) => {
      if (!item.first_air_date) return false;

      return (
        Number(item.first_air_date.slice(0, 4)) ===
        Number(year)
      );
    });

    if (exact) {
      show = exact;
    }
  }

  console.log(
    `  -> ${show.name} (${show.first_air_date || "?"}), TMDB=${show.id}`
  );

  return show;
}

async function getShowDetails(tmdbId) {
  const response = await tmdb.get(`/tv/${tmdbId}`);
  return response.data;
}

async function getTvExternalIds(tmdbId) {
  const response = await tmdb.get(
    `/tv/${tmdbId}/external_ids`
  );

  return response.data;
}

async function getSeason(tmdbId, seasonNumber) {
  const response = await tmdb.get(
    `/tv/${tmdbId}/season/${seasonNumber}`
  );

  return response.data;
}

// =====================================================
// MAIN
// =====================================================

async function main() {
  if (
    !TMDB_API_KEY ||
    TMDB_API_KEY === "ВСТАВЬ_СЮДА_СВОЙ_TMDB_API_KEY"
  ) {
    throw new Error(
      "Insert your TMDB API key into TMDB_API_KEY."
    );
  }

  if (!fs.existsSync(INPUT_FILE)) {
    throw new Error(`File not found: ${INPUT_FILE}`);
  }

  const workbook = XLSX.readFile(INPUT_FILE, {
    cellDates: true,
  });

  if (!workbook.Sheets["Watched"]) {
    throw new Error('Sheet "Watched" not found.');
  }

  if (!workbook.Sheets["Shows"]) {
    throw new Error('Sheet "Shows" not found.');
  }

  const watchedRows = XLSX.utils.sheet_to_json(
    workbook.Sheets["Watched"],
    {
      raw: true,
      defval: "",
    }
  );

  const showRows = XLSX.utils.sheet_to_json(
    workbook.Sheets["Shows"],
    {
      raw: true,
      defval: "",
    }
  );

  const movies = watchedRows.filter((row) =>
    normalizeTitle(row.Title)
  );

  const shows = showRows.filter((row) =>
    normalizeTitle(row.Title)
  );

  console.log("");
  console.log("==============================");
  console.log("MUST EXPORT");
  console.log("==============================");
  console.log(`Movies: ${movies.length}`);
  console.log(`Shows:  ${shows.length}`);
  console.log("");

  const watchedMovies = [];
  const watchedShows = [];
  const watchHistory = [];

  // =====================================================
  // MOVIES
  // =====================================================

  for (let i = 0; i < movies.length; i++) {
    const row = movies[i];

    const title = normalizeTitle(row.Title);
    const year = getYear(row["Release date"]);
    const watchedAt = toISOString(row.Modified);

    console.log(
      `[MOVIE ${i + 1}/${movies.length}] ${title}`
    );

    const movie = await findMovie(title, year);

    const resolvedYear = movie.release_date
      ? Number(movie.release_date.slice(0, 4))
      : year;

    watchedMovies.push({
      plays: 1,
      last_watched_at: watchedAt,
      last_updated_at: watchedAt,

      movie: {
        title: movie.title,
        year: resolvedYear || null,

        ids: {
          trakt: null,
          slug: null,
          imdb: null,
          tmdb: Number(movie.id),
        },
      },
    });

    watchHistory.push({
      id: null,
      watched_at: watchedAt,
      action: "watch",
      type: "movie",

      movie: {
        title: movie.title,
        year: resolvedYear || null,

        ids: {
          trakt: null,
          slug: null,
          imdb: null,
          tmdb: Number(movie.id),
        },
      },
    });

    console.log("  Added.");

    await sleep(100);
  }

  // =====================================================
  // SHOWS
  // =====================================================

  for (let i = 0; i < shows.length; i++) {
    const row = shows[i];

    const title = normalizeTitle(row.Title);
    const year = getYear(row["Release date"]);
    const watchedCount = Number(row.Watched) || 0;
    const watchedAt = toISOString(row.Modified);

    console.log("");
    console.log(
      `[SHOW ${i + 1}/${shows.length}] ${title} — ${watchedCount} watched`
    );

    const show = await findShow(title, year);

    const showDetails = await getShowDetails(show.id);
    const externalIds = await getTvExternalIds(show.id);

    const resolvedYear = show.first_air_date
      ? Number(show.first_air_date.slice(0, 4))
      : year;

    const allEpisodes = [];

    const seasons = (showDetails.seasons || [])
      .filter((season) => season.season_number > 0)
      .sort(
        (a, b) =>
          a.season_number - b.season_number
      );

    for (const season of seasons) {
      const seasonData = await getSeason(
        show.id,
        season.season_number
      );

      const episodes = (
        seasonData.episodes || []
      ).sort(
        (a, b) =>
          a.episode_number - b.episode_number
      );

      for (const episode of episodes) {
        allEpisodes.push({
          season_number: season.season_number,
          episode_number: episode.episode_number,
          tmdb_id: episode.id,
          air_date: episode.air_date || null,
          name: episode.name || "",
        });
      }

      await sleep(100);
    }

    if (watchedCount > allEpisodes.length) {
      throw new Error(
        `${title}: Must says ${watchedCount} watched episodes, but TMDB has ${allEpisodes.length}.`
      );
    }

    const selectedEpisodes =
      allEpisodes.slice(0, watchedCount);

    const seasonsMap = new Map();

    for (const episode of selectedEpisodes) {
      if (!seasonsMap.has(episode.season_number)) {
        seasonsMap.set(episode.season_number, {
          number: episode.season_number,
          episodes: [],
        });
      }

      seasonsMap
        .get(episode.season_number)
        .episodes.push({
          number: episode.episode_number,
          plays: 1,
          last_watched_at: watchedAt,
        });

      watchHistory.push({
        id: null,
        watched_at: watchedAt,
        action: "watch",
        type: "episode",

        episode: {
          season: episode.season_number,
          number: episode.episode_number,

          ids: {
            trakt: null,
            tvdb: null,
            imdb: null,
            tmdb: Number(episode.tmdb_id),
          },
        },

        show: {
          title: show.name,
          year: resolvedYear || null,

          ids: {
            trakt: null,
            slug: null,
            tvdb: externalIds.tvdb_id
              ? Number(externalIds.tvdb_id)
              : null,
            imdb: externalIds.imdb_id || null,
            tmdb: Number(show.id),
          },
        },
      });

      console.log(
        `  S${episode.season_number}E${episode.episode_number} ${episode.name}`
      );
    }

    watchedShows.push({
      plays: selectedEpisodes.length,
      last_watched_at: watchedAt,
      last_updated_at: watchedAt,
      reset_at: null,

      show: {
        title: show.name,
        year: resolvedYear || null,

        ids: {
          trakt: null,
          slug: null,
          tvdb: externalIds.tvdb_id
            ? Number(externalIds.tvdb_id)
            : null,
          imdb: externalIds.imdb_id || null,
          tmdb: Number(show.id),
        },
      },

      seasons: Array.from(
        seasonsMap.values()
      ).sort(
        (a, b) =>
          a.number - b.number
      ),
    });

    console.log(
      `  Added ${selectedEpisodes.length} episodes.`
    );

    await sleep(100);
  }

  // =====================================================
  // CREATE FILES
  // =====================================================

  fs.rmSync(OUT_DIR, {
    recursive: true,
    force: true,
  });

  const watchedDir = path.join(
    OUT_DIR,
    "watched"
  );

  fs.mkdirSync(watchedDir, {
    recursive: true,
  });

  fs.writeFileSync(
    path.join(
      watchedDir,
      "watched-movies.json"
    ),
    JSON.stringify(
      watchedMovies,
      null,
      2
    ),
    "utf8"
  );

  fs.writeFileSync(
    path.join(
      watchedDir,
      "watched-shows.json"
    ),
    JSON.stringify(
      watchedShows,
      null,
      2
    ),
    "utf8"
  );

  fs.writeFileSync(
    path.join(
      watchedDir,
      "watched-history-1.json"
    ),
    JSON.stringify(
      watchHistory,
      null,
      2
    ),
    "utf8"
  );

  // =====================================================
  // ZIP
  // =====================================================

  if (fs.existsSync(OUT_ZIP)) {
    fs.unlinkSync(OUT_ZIP);
  }

  const zip = new AdmZip();

  zip.addLocalFolder(
    watchedDir,
    "watched"
  );

  zip.writeZip(OUT_ZIP);

  console.log("");
  console.log("==============================");
  console.log("DONE");
  console.log("==============================");
  console.log(
    `Movies:         ${watchedMovies.length}`
  );
  console.log(
    `Shows:          ${watchedShows.length}`
  );
  console.log(
    `History events: ${watchHistory.length}`
  );
  console.log("");
  console.log(`Created: ${OUT_ZIP}`);
}

main().catch((error) => {
  console.error("");
  console.error("==============================");
  console.error("FATAL ERROR");
  console.error("==============================");

  console.error(
    error.response?.data ||
      error.stack ||
      error.message ||
      error
  );

  process.exit(1);
});