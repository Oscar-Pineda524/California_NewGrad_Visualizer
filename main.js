// California New Grad Affordability Dashboard
// ------------------------------------------------------------
// Data goal:
// 1. Load household burden, wage, employment, rent, and California county shapes.
// 2. Clean each row to a shared county_id, preferably 5-digit county FIPS.
// 3. Join all datasets onto each county map feature.
// 4. Warn in the console when a county cannot be matched.

const width = 760;
const height = 760;

// Change these paths if your folders or filenames are different.
// Browser D3 loads CSV/JSON directly. SheetJS loads the local OEWS .xlsx files.
const DATA_PATHS = {
  countyTopology: "https://cdn.jsdelivr.net/npm/us-atlas@3/counties-10m.json",
  householdBurden: "Data/household_burden/annual.csv",
  householdLookup: "Data/household_burden/california_counties_burdened_households_series_ids.csv",
  medianRent: "Data/median_gross_rent/2024MedianGrossRent.csv",
  oewsRoot: "Data/Industry_employmet&wages",
  oewsLatestCompact: "Data/Industry_employmet&wages/oews_latest_compact.json",
  oewsLatest: "Data/Industry_employmet&wages/oews_latest.json",
  oewsNormalized: "Data/Industry_employmet&wages/oews_normalized.json"
};

// Browsers cannot list local folders, so this manifest tells D3 which OEWS
// workbooks are available. These are the latest five local OEWS years.
const OEWS_WORKBOOKS_BY_YEAR = {
  2022: [
    "CA-OEWS-Anaheim-Santa Ana-Irvine MD-2022.xlsx",
    "CA-OEWS-Bakersfield MSA-2022.xlsx",
    "CA-OEWS-Chico MSA-2022.xlsx",
    "CA-OEWS-Eastern Sierra-Mother Lode Region-2022.xlsx",
    "CA-OEWS-El Centro MSA-2022.xlsx",
    "CA-OEWS-Fresno MSA-2022.xlsx",
    "CA-OEWS-Hanford-Corcoran MSA-2022.xlsx",
    "CA-OEWS-Los Angeles-Long Beach-Glendale MD-2022.xlsx",
    "CA-OEWS-Madera MSA-2022.xlsx",
    "CA-OEWS-Merced MSA-2022.xlsx",
    "CA-OEWS-Modesto MSA-2022.xlsx",
    "CA-OEWS-Napa MSA-2022.xlsx",
    "CA-OEWS-North Coast Region-2022.xlsx",
    "CA-OEWS-North Valley-Northern Mountains Region-2022.xlsx",
    "CA-OEWS-Oakland-Hayward-Berkeley MD-2022.xlsx",
    "CA-OEWS-Oxnard-Thousand Oaks-Ventura MSA-2022.xlsx",
    "CA-OEWS-Redding MSA-2022.xlsx",
    "CA-OEWS-Riverside-San Bernardino-Ontario MSA-2022.xlsx",
    "CA-OEWS-Sacramento-Roseville-Arden Arcade MSA-2022.xlsx",
    "CA-OEWS-Salinas MSA-2022.xlsx",
    "CA-OEWS-San Diego-Carlsbad MSA-2022.xlsx",
    "CA-OEWS-San Francisco-Redwood City-South San Francisco MD-2022.xlsx",
    "CA-OEWS-San Jose-Sunnyvale-Santa Clara MSA-2022.xlsx",
    "CA-OEWS-San Luis Obispo-Paso Robles-Arroyo Grande MSA-2022.xlsx",
    "CA-OEWS-San Rafael MD-2022.xlsx",
    "CA-OEWS-Santa Cruz-Watsonville MSA-2022.xlsx",
    "CA-OEWS-Santa Maria-Santa Barbara MSA-2022.xlsx",
    "CA-OEWS-Santa Rosa MSA-2022.xlsx",
    "CA-OEWS-Stockton-Lodi MSA-2022.xlsx",
    "CA-OEWS-Vallejo-Fairfield MSA-2022.xlsx",
    "CA-OEWS-Visalia-Porterville MSA-2022.xlsx",
    "CA-OEWS-Yuba City MSA-2022.xlsx"
  ],
  2023: [
    "CA-OEWS-Anaheim-Santa Ana-Irvine MD-2023.xlsx",
    "CA-OEWS-Bakersfield MSA-2023.xlsx",
    "CA-OEWS-Chico MSA-2023.xlsx",
    "CA-OEWS-Eastern Sierra-Mother Lode Region-2023.xlsx",
    "CA-OEWS-El Centro MSA-2023.xlsx",
    "CA-OEWS-Fresno MSA-2023.xlsx",
    "CA-OEWS-Hanford-Corcoran MSA-2023.xlsx",
    "CA-OEWS-Los Angeles-Long Beach-Glendale MD-2023.xlsx",
    "CA-OEWS-Madera MSA-2023.xlsx",
    "CA-OEWS-Merced MSA-2023.xlsx",
    "CA-OEWS-Modesto MSA-2023.xlsx",
    "CA-OEWS-Napa MSA-2023.xlsx",
    "CA-OEWS-North Coast Region-2023.xlsx",
    "CA-OEWS-North Valley-Northern Mountains Region-2023.xlsx",
    "CA-OEWS-Oakland-Hayward-Berkeley MD-2023.xlsx",
    "CA-OEWS-Oxnard-Thousand Oaks-Ventura MSA-2023.xlsx",
    "CA-OEWS-Redding MSA-2023.xlsx",
    "CA-OEWS-Riverside-San Bernardino-Ontario MSA-2023.xlsx",
    "CA-OEWS-Sacramento-Roseville-Arden Arcade MSA-2023.xlsx",
    "CA-OEWS-Salinas MSA-2023.xlsx",
    "CA-OEWS-San Diego-Carlsbad MSA-2023.xlsx",
    "CA-OEWS-San Francisco-Redwood City-South San Francisco MD-2023.xlsx",
    "CA-OEWS-San Jose-Sunnyvale-Santa Clara MSA-2023.xlsx",
    "CA-OEWS-San Luis Obispo-Paso Robles-Arroyo Grande MSA-2023.xlsx",
    "CA-OEWS-San Rafael MD-2023.xlsx",
    "CA-OEWS-Santa Cruz-Watsonville MSA-2023.xlsx",
    "CA-OEWS-Santa Maria-Santa Barbara MSA-2023.xlsx",
    "CA-OEWS-Santa Rosa MSA-2023.xlsx",
    "CA-OEWS-Stockton-Lodi MSA-2023.xlsx",
    "CA-OEWS-Vallejo-Fairfield MSA-2023.xlsx",
    "CA-OEWS-Visalia-Porterville MSA-2023.xlsx",
    "CA-OEWS-Yuba City MSA-2023.xlsx"
  ],
  2024: [
    "CA-OEWS-Anaheim-Santa Ana-Irvine MD-2024.xlsx",
    "CA-OEWS-Bakersfield MSA-2024.xlsx",
    "CA-OEWS-Chico MSA-2024.xlsx",
    "CA-OEWS-Eastern Sierra-Mother Lode Region-2024.xlsx",
    "CA-OEWS-El Centro MSA-2024.xlsx",
    "CA-OEWS-Fresno MSA-2024.xlsx",
    "CA-OEWS-Hanford-Corcoran MSA-2024.xlsx",
    "CA-OEWS-Los Angeles-Long Beach-Glendale MD-2024.xlsx",
    "CA-OEWS-Madera MSA-2024.xlsx",
    "CA-OEWS-Merced MSA-2024.xlsx",
    "CA-OEWS-Modesto MSA-2024.xlsx",
    "CA-OEWS-Napa MSA-2024.xlsx",
    "CA-OEWS-North Coast Region-2024.xlsx",
    "CA-OEWS-North Valley-Northern Mountains Region-2024.xlsx",
    "CA-OEWS-Oakland-Hayward-Berkeley MD-2024.xlsx",
    "CA-OEWS-Oxnard-Thousand Oaks-Ventura MSA-2024.xlsx",
    "CA-OEWS-Redding MSA-2024.xlsx",
    "CA-OEWS-Riverside-San Bernardino-Ontario MSA-2024.xlsx",
    "CA-OEWS-Sacramento-Roseville-Arden Arcade MSA-2024.xlsx",
    "CA-OEWS-Salinas MSA-2024.xlsx",
    "CA-OEWS-San Diego-Carlsbad MSA-2024.xlsx",
    "CA-OEWS-San Francisco-Redwood City-South San Francisco MD-2024.xlsx",
    "CA-OEWS-San Jose-Sunnyvale-Santa Clara MSA-2024.xlsx",
    "CA-OEWS-San Luis Obispo-Paso Robles-Arroyo Grande MSA-2024.xlsx",
    "CA-OEWS-San Rafael MD-2024.xlsx",
    "CA-OEWS-Santa Cruz-Watsonville MSA-2024.xlsx",
    "CA-OEWS-Santa Maria-Santa Barbara MSA-2024.xlsx",
    "CA-OEWS-Santa Rosa MSA-2024.xlsx",
    "CA-OEWS-Stockton-Lodi MSA-2024.xlsx",
    "CA-OEWS-Vallejo-Fairfield MSA-2024.xlsx",
    "CA-OEWS-Visalia-Porterville MSA-2024.xlsx",
    "CA-OEWS-Yuba City MSA-2024.xlsx"
  ],
  2025: [
    "CA-OEWS-Anaheim-Santa Ana-Irvine MD-2025.xlsx",
    "CA-OEWS-Bakersfield-Delano MSA-2025.xlsx",
    "CA-OEWS-Chico MSA-2025.xlsx",
    "CA-OEWS-Eastern Sierra-Mother Lode Region-2025.xlsx",
    "CA-OEWS-El Centro MSA-2025.xlsx",
    "CA-OEWS-Fresno MSA-2025.xlsx",
    "CA-OEWS-Hanford-Corcoran MSA-2025.xlsx",
    "CA-OEWS-Los Angeles-Long Beach-Glendale MD-2025.xlsx",
    "CA-OEWS-Merced MSA-2025.xlsx",
    "CA-OEWS-Modesto MSA-2025.xlsx",
    "CA-OEWS-Napa MSA-2025.xlsx",
    "CA-OEWS-North Coast Region-2025.xlsx",
    "CA-OEWS-North Valley-Northern Mountains Region-2025.xlsx",
    "CA-OEWS-Oakland-Fremont-Berkeley MD-2025.xlsx",
    "CA-OEWS-Oxnard-Thousand Oaks-Ventura MSA-2025.xlsx",
    "CA-OEWS-Redding MSA-2025.xlsx",
    "CA-OEWS-Riverside-San Bernardino-Ontario MSA-2025.xlsx",
    "CA-OEWS-Sacramento-Roseville-Folsom MSA-2025.xlsx",
    "CA-OEWS-Salinas MSA-2025.xlsx",
    "CA-OEWS-San Diego-Chula Vista-Carlsbad MSA-2025.xlsx",
    "CA-OEWS-San Francisco-San Mateo-Redwood City MD-2025.xlsx",
    "CA-OEWS-San Jose-Sunnyvale-Santa Clara MSA-2025.xlsx",
    "CA-OEWS-San Luis Obispo-Paso Robles MSA-2025.xlsx",
    "CA-OEWS-San Rafael MD-2025.xlsx",
    "CA-OEWS-Santa Cruz-Watsonville MSA-2025.xlsx",
    "CA-OEWS-Santa Maria-Santa Barbara MSA-2025.xlsx",
    "CA-OEWS-Santa Rosa-Petaluma MSA-2025.xlsx",
    "CA-OEWS-Stockton-Lodi MSA-2025.xlsx",
    "CA-OEWS-Vallejo MSA-2025.xlsx",
    "CA-OEWS-Visalia MSA-2025.xlsx",
    "CA-OEWS-Yuba City MSA-2025.xlsx"
  ],
  2026: [
    "CA-OEWS-Anaheim-Santa Ana-Irvine MD-2026.xlsx",
    "CA-OEWS-Bakersfield-Delano MSA-2026.xlsx",
    "CA-OEWS-Chico MSA-2026.xlsx",
    "CA-OEWS-Eastern Sierra-Mother Lode Region-2026.xlsx",
    "CA-OEWS-El Centro MSA-2026.xlsx",
    "CA-OEWS-Fresno MSA-2026.xlsx",
    "CA-OEWS-Hanford-Corcoran MSA-2026.xlsx",
    "CA-OEWS-Los Angeles-Long Beach-Glendale MD-2026.xlsx",
    "CA-OEWS-Merced MSA-2026.xlsx",
    "CA-OEWS-Modesto MSA-2026.xlsx",
    "CA-OEWS-Napa MSA-2026.xlsx",
    "CA-OEWS-North Coast Region-2026.xlsx",
    "CA-OEWS-North Valley-Northern Mountains Region-2026.xlsx",
    "CA-OEWS-Oakland-Fremont-Berkeley MD-2026.xlsx",
    "CA-OEWS-Oxnard-Thousand Oaks-Ventura MSA-2026.xlsx",
    "CA-OEWS-Redding MSA-2026.xlsx",
    "CA-OEWS-Riverside-San Bernardino-Ontario MSA-2026.xlsx",
    "CA-OEWS-Sacramento-Roseville-Folsom MSA-2026.xlsx",
    "CA-OEWS-Salinas MSA-2026.xlsx",
    "CA-OEWS-San Diego-Chula Vista-Carlsbad MSA-2026.xlsx",
    "CA-OEWS-San Francisco-San Mateo-Redwood City MD-2026.xlsx",
    "CA-OEWS-San Jose-Sunnyvale-Santa Clara MSA-2026.xlsx",
    "CA-OEWS-San Luis Obispo-Paso Robles MSA-2026.xlsx",
    "CA-OEWS-San Rafael MD-2026.xlsx",
    "CA-OEWS-Santa Cruz-Watsonville MSA-2026.xlsx",
    "CA-OEWS-Santa Maria-Santa Barbara MSA-2026.xlsx",
    "CA-OEWS-Santa Rosa-Petaluma MSA-2026.xlsx",
    "CA-OEWS-Stockton-Lodi MSA-2026.xlsx",
    "CA-OEWS-Vallejo MSA-2026.xlsx",
    "CA-OEWS-Visalia MSA-2026.xlsx",
    "CA-OEWS-Yuba City MSA-2026.xlsx"
  ]
};

// Update these column names if the CSV lookup/rent files change.
const COLUMN_NAMES = {
  householdDate: "observation_date",
  householdLookupCounty: "county",
  householdLookupFips: "county_fips",
  rentGeoId: "GEO_ID",
  rentCountyName: "NAME",
  rentValue: "B25064_001E",
  wageCountyFips: "county_fips",
  wageCountyName: "county",
  // OEWS files are occupation-based. The dashboard uses this as the dropdown category.
  wageIndustry: "industry",
  wageAnnual: "annual_mean_wage",
  employmentCountyFips: "county_fips",
  employmentCountyName: "county",
  employmentIndustry: "industry",
  employmentCount: "employment"
};

// Update these labels if a future OEWS workbook renames its header columns.
const OEWS_REQUIRED_HEADERS = {
  soc: "SOC Code",
  occupation: "Occupational Title",
  employment: "Employment Estimates",
  annualWage: "Mean Annual Wage"
};

const occupationGroupSelect = d3.select("#occupation-group-select");
const detailPanel = d3.select("#detail-panel");
const svg = d3.select("#county-map")
  .attr("viewBox", `0 0 ${width} ${height}`)
  .attr("preserveAspectRatio", "xMidYMid meet");

const mapGroup = svg.append("g");
const legend = d3.select("#legend");

const formatDollars = d3.format("$,.0f");
const formatPercent = d3.format(".0%");
const formatNumber = d3.format(",");

const colorScale = d3.scaleThreshold()
  .domain([0.25, 0.30, 0.35, 0.40])
  .range(["#1f9d6a", "#75c76f", "#ffd166", "#f28e4b", "#c2413b"]);

const affordabilityLabels = [
  "Very affordable",
  "Affordable",
  "Manageable",
  "Stretched",
  "Cost burdened"
];

let counties = [];
let selectedCounty = null;
let joinedDataByCountyId = new Map();
let countyFipsByName = new Map();
let countyNameByFips = new Map();
let normalizedOewsRows = [];
let latestOewsYear = null;
let latestOewsRows = [];
let latestOewsRowsByCountySoc = new Map();

initDashboard();
drawLegend();

occupationGroupSelect.on("change", () => {
  refreshAffordabilityView();
});

function refreshAffordabilityView() {
  updateMap();
  if (selectedCounty) {
    renderCountyDetails(selectedCounty);
  }
}

async function initDashboard() {
  try {
    const rawData = await loadData();
    const cleanedData = cleanAndIndexData(rawData);

    joinedDataByCountyId = cleanedData.joinedDataByCountyId;
    countyFipsByName = cleanedData.countyFipsByName;
    countyNameByFips = cleanedData.countyNameByFips;
    counties = attachDataToCountyFeatures(rawData.countyTopology, joinedDataByCountyId);

    console.log(`Prepared ${joinedDataByCountyId.size} joined counties and ${counties.length} California map features.`);

    updateOccupationGroupDropdown(cleanedData.availableOccupationGroups);
    drawMap();
  } catch (error) {
    d3.select("#map-container")
      .append("p")
      .attr("class", "error-message")
      .text("Data could not be loaded. Check file paths, run a local server, and inspect the console.");
    console.error("Dashboard data loading error:", error);
  }
}

async function loadData() {
  const [
    countyTopology,
    householdBurden,
    householdLookup,
    medianRent
  ] = await Promise.all([
    d3.json(DATA_PATHS.countyTopology),
    d3.csv(DATA_PATHS.householdBurden),
    d3.csv(DATA_PATHS.householdLookup),
    d3.csv(DATA_PATHS.medianRent)
  ]);

  const householdCountyLookup = cleanHouseholdLookup(householdLookup);
  const oewsRows = await loadNormalizedOewsRows(householdCountyLookup);

  console.log(`Loaded datasets: ${householdBurden.length} household rows, ${medianRent.length} rent rows, ${oewsRows.length} OEWS rows.`);

  return {
    countyTopology,
    householdBurden,
    householdLookup,
    medianRent,
    oewsRows
  };
}

async function loadNormalizedOewsRows(householdCountyLookup) {
  try {
    const compactRows = await loadCompactOewsRows(DATA_PATHS.oewsLatestCompact, householdCountyLookup.countyNameByFips);

    if (compactRows.length) {
      console.log(`Loaded ${compactRows.length} compact OEWS rows from ${DATA_PATHS.oewsLatestCompact}.`);
      logOewsYearSummary(compactRows);
      return compactRows;
    }
  } catch (error) {
    console.warn(`Compact OEWS JSON was not found: ${DATA_PATHS.oewsLatestCompact}`, error);
  }

  const preprocessedPaths = [
    DATA_PATHS.oewsLatest,
    DATA_PATHS.oewsNormalized
  ];

  for (const path of preprocessedPaths) {
    try {
      const rows = await d3.json(path);

      if (Array.isArray(rows) && rows.length) {
        console.log(`Loaded ${rows.length} preprocessed OEWS rows from ${path}.`);
        logOewsYearSummary(rows);
        return rows;
      }

      console.warn(`Preprocessed OEWS JSON is empty or invalid: ${path}`);
    } catch (error) {
      console.warn(`Preprocessed OEWS JSON was not found: ${path}`, error);
    }
  }

  console.warn("Falling back to slower browser Excel parsing.");
  const oewsWorkbookPaths = getOewsWorkbookPaths();
  console.log("OEWS workbook paths loaded for fallback Excel parsing:", oewsWorkbookPaths);
  return loadOewsExcelWorkbooks(oewsWorkbookPaths, householdCountyLookup.countyIdByName);
}

async function loadCompactOewsRows(path, countyNameByFips) {
  const compact = await d3.json(path);

  if (!compact?.rows?.length || !compact?.groups?.length) {
    console.warn(`Compact OEWS JSON is empty or invalid: ${path}`, compact);
    return [];
  }

  return expandCompactOewsRows(compact, countyNameByFips, path);
}

function expandCompactOewsRows(compact, countyNameByFips, sourceFile) {
  const groupTitleByCode = new Map(compact.groups);

  return compact.rows.map(([countyId, socCode, employment, meanAnnualWage]) => {
    const occupationTitle = groupTitleByCode.get(socCode) || socCode;

    return {
      year: compact.year,
      county: countyNameByFips.get(countyId) || "",
      county_id: countyId,
      soc_code: socCode,
      occupation_title: occupationTitle,
      occupation_group: occupationTitle,
      employment,
      mean_annual_wage: meanAnnualWage,
      source_file: sourceFile,
      occupation_level: "group"
    };
  });
}

function getOewsWorkbookPaths() {
  return Object.entries(OEWS_WORKBOOKS_BY_YEAR)
    .flatMap(([year, fileNames]) => fileNames.map((fileName) => ({
      year: Number(year),
      path: `${DATA_PATHS.oewsRoot}/all-oews-${year}/${fileName}`
    })));
}

async function loadOewsExcelWorkbooks(workbooks, countyIdByName) {
  if (!window.XLSX) {
    throw new Error("SheetJS XLSX library is missing. Check the xlsx.full.min.js script tag in index.html.");
  }

  const workbookRows = await Promise.all(
    workbooks.map((workbook) => loadOewsExcelWorkbook(workbook, countyIdByName))
  );
  const rows = workbookRows.flat();
  console.log(`Parsed ${rows.length} normalized OEWS rows across ${workbooks.length} workbooks.`);
  logOewsYearSummary(rows);
  return rows;
}

async function loadOewsExcelWorkbook(workbookInfo, countyIdByName) {
  const { year, path } = workbookInfo;

  try {
    const response = await fetch(encodeURI(path));

    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText}`);
    }

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheetName = getOewsSheetName(workbook);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: ""
    });

    const parsedRows = parseOewsWorksheetRows(rows, {
      year,
      path,
      countyIdByName
    });
    console.log(`Parsed ${parsedRows.length} normalized OEWS rows from ${path}`);
    return parsedRows;
  } catch (error) {
    console.warn(`Could not load OEWS Excel workbook "${path}".`, error);
    return [];
  }
}

function getOewsSheetName(workbook) {
  return workbook.SheetNames.find((sheetName) => sheetName.trim() === "OEWS Data")
    || workbook.SheetNames.find((sheetName) => /oews?data/i.test(sheetName))
    || workbook.SheetNames[0];
}

function parseOewsWorksheetRows(rows, options) {
  const { year, path, countyIdByName } = options;
  const counties = getOewsWorkbookCounties(rows);
  const headerIndex = rows.findIndex((row) => (
    row.includes(OEWS_REQUIRED_HEADERS.soc)
    && row.includes(OEWS_REQUIRED_HEADERS.occupation)
  ));

  if (headerIndex === -1) {
    console.warn(`OEWS workbook is missing the expected header row: ${path}`);
    return [];
  }

  if (!counties.length) {
    console.warn(`OEWS workbook has no county list and will be skipped for county joins: ${path}`);
    return [];
  }

  const countyMatches = counties
    .map((county) => ({
      county,
      county_id: countyIdByName.get(normalizeCountyName(county)) || null
    }))
    .filter((county) => {
      if (!county.county_id) {
        console.warn(`OEWS county failed to match lookup: "${county.county}" in ${path}`);
        return false;
      }
      return true;
    });

  if (!countyMatches.length) {
    console.warn(`No OEWS counties matched lookup for ${path}`);
    return [];
  }

  const headers = rows[headerIndex].map((header) => String(header).trim());
  const indexes = {
    soc: headers.indexOf(OEWS_REQUIRED_HEADERS.soc),
    occupation: headers.indexOf(OEWS_REQUIRED_HEADERS.occupation),
    employment: findHeaderIndex(headers, OEWS_REQUIRED_HEADERS.employment),
    annualWage: findHeaderIndex(headers, OEWS_REQUIRED_HEADERS.annualWage)
  };

  if (indexes.soc === -1 || indexes.occupation === -1 || indexes.employment === -1 || indexes.annualWage === -1) {
    console.warn(`OEWS workbook has unexpected columns. Update parseOewsWorksheetRows(): ${path}`, headers);
    return [];
  }

  const dataRows = rows.slice(headerIndex + 1)
    .map((row) => toOewsOccupationRow(row, indexes))
    .filter((row) => row && getOccupationLevel(row.soc_code) === "group");
  const groupTitleByPrefix = getOccupationGroupLookup(dataRows);

  return dataRows.flatMap((row) => {
    const occupationGroup = getOccupationGroup(row, groupTitleByPrefix);

    return countyMatches.map((county) => ({
      year,
      county: county.county,
      county_id: county.county_id,
      soc_code: row.soc_code,
      occupation_title: row.occupation_title,
      occupation_group: occupationGroup,
      employment: row.employment,
      mean_annual_wage: row.mean_annual_wage,
      source_file: path,
      occupation_level: getOccupationLevel(row.soc_code)
    }));
  });
}

function toOewsOccupationRow(row, indexes) {
  const socCode = String(row[indexes.soc] || "").trim();
  const occupationTitle = cleanIndustryName(row[indexes.occupation]);

  if (!isSocCode(socCode) || !occupationTitle) {
    return null;
  }

  return {
    soc_code: socCode,
    occupation_title: occupationTitle,
    employment: parseNumber(row[indexes.employment]),
    mean_annual_wage: parseNumber(row[indexes.annualWage])
  };
}

function getOccupationGroupLookup(rows) {
  const groupTitleByPrefix = new Map();

  rows.forEach((row) => {
    if (row.soc_code.endsWith("-0000") && row.soc_code !== "00-0000") {
      groupTitleByPrefix.set(getSocPrefix(row.soc_code), row.occupation_title);
    }
  });

  return groupTitleByPrefix;
}

function getOccupationGroup(row, groupTitleByPrefix) {
  if (row.soc_code === "00-0000") return "Total all occupations";
  return groupTitleByPrefix.get(getSocPrefix(row.soc_code)) || row.occupation_title;
}

function getOccupationGroups(data) {
  const groupRowsByCode = new Map();

  data.forEach((row) => {
    if (row.occupation_level !== "group") return;

    const groupCode = getSocGroupCode(row.soc_code);
    if (!groupCode) return;

    const existing = groupRowsByCode.get(groupCode) || {
      groupCode,
      groupTitle: row.occupation_title,
      occupationCount: 0
    };

    existing.groupTitle = existing.groupTitle || row.occupation_title;
    groupRowsByCode.set(groupCode, existing);
  });

  return Array.from(groupRowsByCode.values())
    .sort((a, b) => d3.ascending(a.groupCode, b.groupCode));
}

function getSelectedOccupationRows(data, groupCode, occupationCode) {
  const normalizedGroupCode = getSocGroupCode(groupCode);

  if (!normalizedGroupCode) {
    console.warn(`Cannot select occupation rows for invalid SOC group code: ${groupCode}`);
    return [];
  }

  return data.filter((row) => row.soc_code === normalizedGroupCode);
}

function indexOewsRowsByCountySoc(rows) {
  const rowsByCountySoc = new Map();

  rows.forEach((row) => {
    if (!row.county_id || !row.soc_code) return;
    rowsByCountySoc.set(makeCountySocKey(row.county_id, row.soc_code), row);
  });

  return rowsByCountySoc;
}

function getOewsCountySocRow(countyId, socCode) {
  if (!countyId || !socCode) return null;
  return latestOewsRowsByCountySoc.get(makeCountySocKey(countyId, socCode)) || null;
}

function getOccupationLevel(socCode) {
  if (socCode === "00-0000") return "total";
  if (socCode.endsWith("-0000")) return "group";
  return "occupation";
}

function getSocGroupCode(socCode) {
  if (!isSocCode(socCode)) return null;

  const prefix = getSocPrefix(socCode);
  if (prefix === "00") return null;
  return `${prefix}-0000`;
}

function getSocPrefix(socCode) {
  return String(socCode).slice(0, 2);
}

function isSocCode(value) {
  return /^\d{2}-\d{4}$/.test(String(value || ""));
}

function logOewsYearSummary(rows) {
  const summary = d3.rollups(
    rows,
    (yearRows) => ({
      rows: yearRows.length,
      counties: new Set(yearRows.map((row) => row.county_id)).size,
      occupationGroups: new Set(yearRows.map((row) => row.soc_code)).size
    }),
    (row) => row.year
  );

  console.log("Normalized OEWS rows by year:", summary);
}

function getOewsWorkbookCounties(rows) {
  const countyRow = rows.find((row) => String(row[0] || "").startsWith("Counties:"));

  if (!countyRow) return [];

  return String(countyRow[0])
    .replace(/^Counties:\s*/i, "")
    .replace(/\s*,?\s+and\s+/gi, ", ")
    .split(/\s*,\s*/)
    .map((countyName) => cleanCountyName(countyName))
    .filter(Boolean);
}

function findHeaderIndex(headers, text) {
  return headers.findIndex((header) => header.toLowerCase().includes(text.toLowerCase()));
}

function cleanAndIndexData(rawData) {
  const lookup = cleanHouseholdLookup(rawData.householdLookup);
  const householdBurdenByCountyId = cleanHouseholdBurden(rawData.householdBurden, lookup.countyIdByName);
  const rentByCountyId = cleanMedianRent(rawData.medianRent, lookup.countyIdByName);
  normalizedOewsRows = rawData.oewsRows;

  latestOewsYear = d3.max(normalizedOewsRows, (row) => row.year);
  latestOewsRows = normalizedOewsRows.filter((row) => row.year === latestOewsYear);
  latestOewsRowsByCountySoc = indexOewsRowsByCountySoc(latestOewsRows);

  console.log(`Cleaned ${householdBurdenByCountyId.size} household burden counties and ${rentByCountyId.size} rent counties.`);
  console.log(`Using ${latestOewsRows.length} OEWS rows from ${latestOewsYear} for occupation affordability.`);
  console.log(`Indexed ${latestOewsRowsByCountySoc.size} county/SOC wage rows.`);

  const joinedDataByCountyId = joinCountyDatasets({
    countyNameByFips: lookup.countyNameByFips,
    householdBurdenByCountyId,
    rentByCountyId,
    wagesByCountyIndustry: new Map(),
    employmentByCountyIndustry: new Map()
  });

  return {
    joinedDataByCountyId,
    availableOccupationGroups: getOccupationGroups(latestOewsRows),
    countyFipsByName: lookup.countyIdByName,
    countyNameByFips: lookup.countyNameByFips,
    normalizedOewsRows
  };
}

function cleanHouseholdLookup(rows) {
  const countyIdByName = new Map();
  const countyNameByFips = new Map();

  rows.forEach((row) => {
    const countyName = cleanCountyName(row[COLUMN_NAMES.householdLookupCounty]);
    const countyId = cleanCountyId(row[COLUMN_NAMES.householdLookupFips]);

    if (!countyName || !countyId) {
      console.warn("Household lookup row is missing county name or FIPS:", row);
      return;
    }

    countyIdByName.set(normalizeCountyName(countyName), countyId);
    countyNameByFips.set(countyId, countyName);
  });

  return { countyIdByName, countyNameByFips };
}

function cleanHouseholdBurden(rows, countyIdByName) {
  const latestRow = getLatestDatedRow(rows, COLUMN_NAMES.householdDate);
  const householdBurdenByCountyId = new Map();

  if (!latestRow) {
    console.warn("No household burden rows found.");
    return householdBurdenByCountyId;
  }

  Object.entries(latestRow).forEach(([columnName, value]) => {
    if (columnName === COLUMN_NAMES.householdDate) return;

    const countyName = cleanCountyName(columnName);
    const countyId = countyIdByName.get(normalizeCountyName(countyName));

    if (!countyId) {
      console.warn(`Household burden county failed to match: "${countyName}"`);
      return;
    }

    householdBurdenByCountyId.set(countyId, {
      county_id: countyId,
      county: countyName,
      year: getYear(latestRow[COLUMN_NAMES.householdDate]),
      burdenedHouseholdsPct: parseNumber(value)
    });
  });

  return householdBurdenByCountyId;
}

function cleanMedianRent(rows, countyIdByName) {
  const rentByCountyId = new Map();

  rows.forEach((row) => {
    // Census files often include a second metadata row where GEO_ID is "Geography".
    if (row[COLUMN_NAMES.rentGeoId] === "Geography") return;

    const countyName = cleanCountyName(row[COLUMN_NAMES.rentCountyName]);
    const countyIdFromGeoId = extractCountyFips(row[COLUMN_NAMES.rentGeoId]);
    const countyId = countyIdFromGeoId || countyIdByName.get(normalizeCountyName(countyName));

    if (!countyId) {
      console.warn("Median rent county failed to match:", row);
      return;
    }

    rentByCountyId.set(countyId, {
      county_id: countyId,
      county: countyName,
      medianGrossRent: parseNumber(row[COLUMN_NAMES.rentValue])
    });
  });

  return rentByCountyId;
}

function cleanIndustryRows(rows, config) {
  const rowsByCountyIndustry = new Map();

  rows.forEach((row) => {
    const countyName = cleanCountyName(row[config.countyNameColumn]);
    const countyId = cleanCountyId(row[config.countyFipsColumn])
      || config.countyIdByName.get(normalizeCountyName(countyName));
    const industry = cleanIndustryName(row[config.industryColumn]);
    const value = parseNumber(row[config.valueColumn]);

    // Change config.* column names above if you see these warnings.
    if (!countyId || !industry) {
      console.warn(`Industry row failed to match for ${config.valueName}:`, row);
      return;
    }

    const key = makeCountyIndustryKey(countyId, industry);
    const existing = rowsByCountyIndustry.get(key) || {
      county_id: countyId,
      county: countyName,
      industry
    };

    existing[config.valueName] = value;
    rowsByCountyIndustry.set(key, existing);
  });

  return rowsByCountyIndustry;
}

function joinCountyDatasets(data) {
  const joined = new Map();

  data.countyNameByFips.forEach((countyName, countyId) => {
    joined.set(countyId, {
      county_id: countyId,
      county: countyName,
      householdBurden: null,
      rent: null,
      industries: new Map()
    });
  });

  data.householdBurdenByCountyId.forEach((householdBurden, countyId) => {
    getOrCreateCounty(joined, countyId, householdBurden.county).householdBurden = householdBurden;
  });

  data.rentByCountyId.forEach((rent, countyId) => {
    getOrCreateCounty(joined, countyId, rent.county).rent = rent;
  });

  data.wagesByCountyIndustry.forEach((wageRow) => {
    const county = getOrCreateCounty(joined, wageRow.county_id, wageRow.county);
    const industry = getOrCreateIndustry(county, wageRow.industry);
    industry.annualWage = wageRow.annualWage;
  });

  data.employmentByCountyIndustry.forEach((employmentRow) => {
    const county = getOrCreateCounty(joined, employmentRow.county_id, employmentRow.county);
    const industry = getOrCreateIndustry(county, employmentRow.industry);
    industry.employment = employmentRow.employment;
  });

  joined.forEach((county, countyId) => {
    if (!county.householdBurden) {
      console.warn(`No household burden match for ${county.county} (${countyId})`);
    }

    if (!county.rent) {
      console.warn(`No median rent match for ${county.county} (${countyId})`);
    }
  });

  return joined;
}

function attachDataToCountyFeatures(countyTopology, dataByCountyId) {
  const features = topojson
    .feature(countyTopology, countyTopology.objects.counties)
    .features
    .filter((feature) => getFeatureCountyId(feature)?.startsWith("06"));

  return features.map((feature) => {
    const countyId = getFeatureCountyId(feature);
    const joinedData = dataByCountyId.get(countyId);

    if (!countyId) {
      console.warn("Map feature is missing a usable county_id:", feature.properties);
    } else if (!joinedData) {
      console.warn(`Map county failed to match joined data: ${getFeatureCountyName(feature)} (${countyId})`);
    }

    feature.properties.county_id = countyId;
    feature.properties.affordabilityData = joinedData || null;
    return feature;
  });
}

function drawMap() {
  const projection = d3.geoMercator()
    .fitSize([width, height], {
      type: "FeatureCollection",
      features: counties
    });

  const path = d3.geoPath(projection);

  mapGroup.selectAll("path")
    .data(counties)
    .join("path")
    .attr("class", "county")
    .attr("d", path)
    .attr("tabindex", 0)
    .attr("role", "button")
    .attr("aria-label", (county) => getCountyName(county))
    .on("mouseenter focus", function(event, county) {
      d3.select(this).classed("is-hovered", true);
      renderCountyDetails(county);
    })
    .on("mouseleave blur", function() {
      d3.select(this).classed("is-hovered", false);
      if (selectedCounty) {
        renderCountyDetails(selectedCounty);
      }
    })
    .on("click", function(event, county) {
      selectedCounty = county;
      mapGroup.selectAll(".county").classed("is-selected", false);
      d3.select(this).classed("is-selected", true);
      renderCountyDetails(county);
    })
    .on("keydown", function(event, county) {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectedCounty = county;
        mapGroup.selectAll(".county").classed("is-selected", false);
        d3.select(this).classed("is-selected", true);
        renderCountyDetails(county);
      }
    });

  updateMap();
}

function updateMap() {
  mapGroup.selectAll(".county")
    .transition()
    .duration(350)
    .attr("fill", (county) => {
      const ratio = getHousingCostRatio(county);
      return ratio == null ? "#d8d4c9" : colorScale(ratio);
    });
}

function drawLegend() {
  const thresholds = ["< 25%", "25-30%", "30-35%", "35-40%", "> 40%"];

  const legendItems = legend.selectAll(".legend-item")
    .data(colorScale.range())
    .join("div")
    .attr("class", "legend-item");

  legendItems.append("span")
    .attr("class", "legend-swatch")
    .style("background-color", (color) => color);

  legendItems.append("span")
    .text((color, index) => `${thresholds[index]} ${affordabilityLabels[index]}`);
}

function renderCountyDetails(county) {
  const countyData = county.properties.affordabilityData;
  const occupationData = getSelectedOccupationData(county);
  const selection = getOccupationSelection();
  const monthlyRent = getMonthlyRent(county);
  const typicalOccupationWage = getTypicalOccupationWage(county);
  const monthlyIncome = getMonthlyIncome(county);
  const ratio = getHousingCostRatio(county);

  detailPanel.html(`
    <p class="eyebrow">County Details</p>
    <h2>${getCountyName(county)}</h2>
    <dl>
      <div>
        <dt>County ID</dt>
        <dd>${county.properties.county_id || "Missing"}</dd>
      </div>
      <div>
        <dt>Median gross rent</dt>
        <dd>${monthlyRent == null ? "No match" : formatDollars(monthlyRent)}</dd>
      </div>
      <div>
        <dt>Selected occupation</dt>
        <dd>${selection.label || "No occupation selected"}</dd>
      </div>
      <div>
        <dt>Typical annual occupation wage</dt>
        <dd>${typicalOccupationWage == null ? "No match" : formatDollars(typicalOccupationWage)}</dd>
      </div>
      <div>
        <dt>Monthly occupation income</dt>
        <dd>${monthlyIncome == null ? "No match" : formatDollars(monthlyIncome)}</dd>
      </div>
      <div>
        <dt>Wage source</dt>
        <dd>${occupationData?.sourceLabel || "No match"}</dd>
      </div>
      <div>
        <dt>Housing cost ratio</dt>
        <dd>${ratio == null ? "No match" : formatPercent(ratio)}</dd>
      </div>
      <div>
        <dt>Affordability category</dt>
        <dd>${getAffordabilityLabel(ratio)}</dd>
      </div>
      <div>
        <dt>Household burden</dt>
        <dd>${countyData?.householdBurden?.burdenedHouseholdsPct == null ? "No match" : `${countyData.householdBurden.burdenedHouseholdsPct}%`}</dd>
      </div>
      <div>
        <dt>Employment</dt>
        <dd>${occupationData?.employment == null ? "No match" : formatNumber(occupationData.employment)}</dd>
      </div>
    </dl>
    <p class="panel-note">
      Formula: median_gross_rent divided by monthly occupation income, where
      monthly occupation income = typical_industry_wage / 12.
    </p>
  `);
}

function updateOccupationGroupDropdown(groups) {
  if (!groups.length) {
    console.warn("No OEWS occupation groups found. Check normalized OEWS rows in the console.");
    occupationGroupSelect
      .property("disabled", true)
      .selectAll("option")
      .data([{ groupCode: "", groupTitle: "No occupation groups found" }])
      .join("option")
      .attr("value", (group) => group.groupCode)
      .text((group) => group.groupTitle);
    return;
  }

  occupationGroupSelect
    .property("disabled", false)
    .selectAll("option")
    .data(groups)
    .join("option")
    .attr("value", (group) => group.groupCode)
    .text((group) => `${group.groupCode} ${group.groupTitle}`);
}

function getMonthlyRent(county) {
  return county.properties.affordabilityData?.rent?.medianGrossRent ?? null;
}

function getMonthlyIncome(county) {
  const typicalOccupationWage = getTypicalOccupationWage(county);

  if (typicalOccupationWage == null) return null;
  return typicalOccupationWage / 12;
}

function getTypicalOccupationWage(county) {
  return getSelectedOccupationData(county)?.annualWage ?? null;
}

function getHousingCostRatio(county) {
  const monthlyRent = getMonthlyRent(county);
  const monthlyIncome = getMonthlyIncome(county);

  if (!monthlyRent || !monthlyIncome) return null;
  return monthlyRent / monthlyIncome;
}

function getSelectedOccupationData(county) {
  const countyId = county.properties.county_id;
  const selection = getOccupationSelection();

  if (!countyId || !selection.groupCode) return null;

  const groupRow = getOewsCountySocRow(countyId, selection.groupCode);

  if (groupRow?.mean_annual_wage != null) {
    return {
      annualWage: groupRow.mean_annual_wage,
      employment: groupRow.employment,
      sourceLabel: "OEWS group-level wage",
      row: groupRow
    };
  }

  console.warn(`No OEWS group wage available for ${getCountyName(county)} (${countyId}), ${selection.groupCode}.`);
  return null;
}

function getOccupationSelection() {
  const groupCode = occupationGroupSelect.property("value");
  const groupLabel = getSelectedOccupationGroupLabel(groupCode);

  return {
    groupCode,
    label: groupLabel
  };
}

function getSelectedOccupationGroupLabel(groupCode) {
  const selectedOptionText = occupationGroupSelect
    .selectAll("option")
    .filter(function() {
      return this.value === groupCode;
    })
    .text();

  return selectedOptionText || groupCode;
}

function getOrCreateCounty(joined, countyId, countyName) {
  if (!joined.has(countyId)) {
    joined.set(countyId, {
      county_id: countyId,
      county: countyName || "Unknown County",
      householdBurden: null,
      rent: null,
      industries: new Map()
    });
  }

  return joined.get(countyId);
}

function getOrCreateIndustry(county, industryName) {
  const industryId = normalizeIndustryName(industryName);

  if (!county.industries.has(industryId)) {
    county.industries.set(industryId, {
      industry: industryName,
      annualWage: null,
      employment: null
    });
  }

  return county.industries.get(industryId);
}

function getLatestDatedRow(rows, dateColumn) {
  return rows
    .filter((row) => row[dateColumn])
    .sort((a, b) => d3.descending(new Date(a[dateColumn]), new Date(b[dateColumn])))[0];
}

function getAvailableIndustries(wagesByCountyIndustry, employmentByCountyIndustry) {
  const industries = new Map();

  wagesByCountyIndustry.forEach((row) => {
    industries.set(normalizeIndustryName(row.industry), row.industry);
  });

  employmentByCountyIndustry.forEach((row) => {
    industries.set(normalizeIndustryName(row.industry), row.industry);
  });

  return Array.from(industries.values()).sort(d3.ascending);
}

function getFeatureCountyId(feature) {
  const props = feature.properties || {};

  if (feature.id) {
    return cleanCountyId(feature.id);
  }

  // Add your GeoJSON's exact county FIPS property here if it uses another name.
  if (props.STATEFP && props.COUNTYFP) {
    return cleanCountyId(`${props.STATEFP}${props.COUNTYFP}`);
  }

  const directFips = cleanCountyId(
    props.county_id
      || props.COUNTY_FIPS
      || props.county_fips
      || props.FIPS
  );

  if (directFips) return directFips;

  // County GeoJSON files often use GEOID = 06001. Place/city GeoJSON files can
  // use 7-digit place GEOIDs, so only trust GEOID when it is exactly 5 digits.
  const geoIdFips = cleanStrictCountyFips(props.GEOID || props.geoid);
  if (geoIdFips) return geoIdFips;

  const countyName = getFeatureCountyName(feature);
  return countyFipsByName.get(normalizeCountyName(countyName)) || null;
}

function getFeatureCountyName(feature) {
  const props = feature.properties || {};

  // Add your GeoJSON's exact county name property here if it uses another name.
  return cleanCountyName(
    props.county
      || props.COUNTY
      || props.name
      || props.NAME
      || props.CountyName
      || props.county_name
  );
}

function getCountyName(feature) {
  const countyId = feature.properties.county_id;
  return countyNameByFips.get(countyId)
    || getFeatureCountyName(feature)
    || "Unknown County";
}

function cleanCountyId(value) {
  if (value == null || value === "") return null;

  const digits = String(value).replace(/\D/g, "");

  // Census GEO_ID values often look like 0500000US06001.
  if (digits.length >= 5) {
    return digits.slice(-5);
  }

  return digits.padStart(5, "0");
}

function cleanStrictCountyFips(value) {
  const digits = String(value || "").replace(/\D/g, "");
  return digits.length === 5 ? digits : null;
}

function extractCountyFips(value) {
  const text = String(value || "");
  const match = text.match(/US(\d{5})$/);
  return match ? match[1] : cleanCountyId(value);
}

function cleanCountyName(value) {
  const countyName = String(value || "")
    .replace(/,\s*California$/i, "")
    .replace(/^and\s+/i, "")
    .trim();

  if (!countyName) return "";
  return /\bcounty$/i.test(countyName) ? countyName : `${countyName} County`;
}

function normalizeCountyName(value) {
  return cleanCountyName(value)
    .toLowerCase()
    .replace(/\s+county$/i, "")
    .replace(/[^a-z0-9]/g, "");
}

function cleanIndustryName(value) {
  return String(value || "").trim();
}

function normalizeIndustryName(value) {
  return cleanIndustryName(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function makeCountyIndustryKey(countyId, industry) {
  return `${countyId}__${normalizeIndustryName(industry)}`;
}

function makeCountySocKey(countyId, socCode) {
  return `${countyId}__${socCode}`;
}

function parseNumber(value) {
  if (value == null || value === "" || value === "-" || value === "**") {
    return null;
  }

  const parsed = Number(String(value).replace(/[$,%\s,]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function getYear(dateText) {
  return String(dateText || "").slice(0, 4);
}

function getAffordabilityLabel(ratio) {
  if (ratio == null) return "No data";
  if (ratio < 0.25) return affordabilityLabels[0];
  if (ratio < 0.30) return affordabilityLabels[1];
  if (ratio < 0.35) return affordabilityLabels[2];
  if (ratio < 0.40) return affordabilityLabels[3];
  return affordabilityLabels[4];
}
