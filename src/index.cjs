const fs = require("node:fs");
const path = require("node:path");
const xl = require("xlsx");

// ==========
// Constants.
// ==========

const dirInput = path.join(__dirname, "../excel");
const fileOutputCsv = "output/linkedin-data-list.csv";
const fileOutputExcel = "output/linkedin-data-list.xlsx";
const regexNumber = /^[0-9,.]+$/;
const sheetNameOne = "PERFORMANCE";
const sheetNameTwo = "TOP DEMOGRAPHICS";

// ====================
// Helper: format date.
// ====================

const formatDate = (str) => {
  // Get date.
  const dateObj = new Date(str);
  const YYYY = dateObj.getFullYear();
  const MM = String(dateObj.getMonth() + 1).padStart(2, "0");
  const DD = String(dateObj.getDate()).padStart(2, "0");
  const YYYY_MM_DD = `${YYYY}-${MM}-${DD}`;

  // Expose string.
  return YYYY_MM_DD;
};

// ====================
// Helper: format time.
// ====================

const formatTime = (str) => {
  // Get time.
  const dateObj = new Date(`January 1, 1970 ${str}`);
  const h = String(dateObj.getHours()).padStart(2, "0");
  const m = String(dateObj.getMinutes()).padStart(2, "0");
  const time = `${h}:${m}`;

  // Expose string.
  return time;
};

// ======================
// Helper: format object.
// ======================

const colOrder = {
  post_date: null,
  post_publish_time: null,
  impressions: 0,
  members_reached: 0,
  reactions: 0,
  comments: 0,
  reposts: 0,
  employees_1_to_10: 0,
  employees_51_to_200: 0,
  employees_201_to_500: 0,
  employees_501_to_1000: 0,
  employees_1001_to_5000: 0,
  employees_5001_to_10000: 0,
  employees_10001_or_more: 0,
  reactions_top_job_title: null,
  reactions_top_location: null,
  reactions_top_industry: null,
  comments_top_job_title: null,
  comments_top_location: null,
  comments_top_industry: null,
  post_url: null,
};

const formatRow = (obj) => {
  // Set later.
  const newObj = {};

  // Loop through.
  for (const [key, fallback] of Object.entries(colOrder)) {
    // Update.
    newObj[key] = obj[key] || fallback;
  }

  // Expose object.
  return newObj;
};

// ============================
// Helper: sort by recent date.
// ============================

const sortByRecentDate = (a, b) => {
  // Expose number.
  return b.post_date.localeCompare(a.post_date);
};

// ========================
// Parse directory & files.
// ========================

fs.readdir(dirInput, (error, fileList) => {
  // Has error: YES.
  if (error) {
    // Log.
    console.error("Error reading directory:", error);

    // Early exit.
    return;
  }

  // Filter for Excel files.
  fileList = fileList.filter((x) => !x.startsWith("~") && x.endsWith(".xlsx"));

  // Files exist: NO.
  if (!fileList.length) {
    // Log.
    console.error(`No Excel files found: ${dirInput}`);

    // Early exit.
    return;
  }

  // Set later.
  const jsonList = [];

  // Loop through.
  fileList.forEach((fileName) => {
    // Get file path.
    const filePath = path.join(dirInput, fileName);

    try {
      // Get file data.
      const fileData = xl.readFile(filePath);

      // Set later.
      const jsonRow = {};

      // Get sheets.
      const sheetOne = fileData?.Sheets?.[sheetNameOne];
      const sheetTwo = fileData?.Sheets?.[sheetNameTwo];

      // Sheet one exists: YES.
      if (sheetOne) {
        // Get sheet data.
        const sheetDataOne = xl.utils.sheet_to_json(sheetOne, { header: 1 });

        // Is array: YES.
        if (Array.isArray(sheetDataOne)) {
          // Loop through.
          sheetDataOne.forEach(([key, value]) => {
            // Clean up.
            key = String(key || "")
              .trim()
              .toLowerCase()
              .replace(/\s+/g, "_");

            value = String(value ?? "")
              .trim()
              .replace(/\s+/g, " ");

            // Is numeric: YES.
            if (regexNumber.test(value)) {
              // Update.
              value = parseFloat(value.replaceAll(",", ""));
            }

            // Get boolean.
            const isValid = key && (value || typeof value === "number");

            // Is valid: NO.
            if (!isValid) {
              // Early exit.
              return;
            }

            // Normalize "comments" key: YES.
            if (["comment", "comments"].includes(key)) {
              // Update.
              jsonRow.comments = value;

              // Early exit.
              return;
            }

            // Normalize "impressions" key: YES.
            if (["impression", "impressions"].includes(key)) {
              // Update.
              jsonRow.impressions = value;

              // Early exit.
              return;
            }

            // Normalize "reactions" key: YES.
            if (["reaction", "reactions"].includes(key)) {
              // Update.
              jsonRow.reactions = value;

              // Early exit.
              return;
            }

            // Normalize "reposts" key: YES.
            if (["repost", "reposts"].includes(key)) {
              // Update.
              jsonRow.reposts = value;

              // Early exit.
              return;
            }

            // Normalize time format: YES.
            if (key === "post_publish_time") {
              // Update.
              jsonRow.post_publish_time = formatTime(value);

              // Early exit.
              return;
            }

            // Normalize "YYYY-MM-DD" format: YES.
            if (key === "post_date") {
              // Update.
              jsonRow.post_date = formatDate(value);

              // Early exit.
              return;
            }

            // Disambiguate "top job title": YES.
            if (key === "top_job_title") {
              // Already exists: NO.
              if (!jsonRow.reactions_top_job_title) {
                // Assume reactions.
                jsonRow.reactions_top_job_title = value;
              } else {
                // Otherwise, assume comments.
                jsonRow.comments_top_job_title = value;
              }

              // Early exit.
              return;
            }

            // Disambiguate "top location": YES.
            if (key === "top_location") {
              // Already exists: NO.
              if (!jsonRow.reactions_top_location) {
                // Assume reactions.
                jsonRow.reactions_top_location = value;
              } else {
                // Otherwise, assume comments.
                jsonRow.comments_top_location = value;
              }

              // Early exit.
              return;
            }

            // Disambiguate "top industry": YES.
            if (key === "top_industry") {
              // Already exists: NO.
              if (!jsonRow.reactions_top_industry) {
                // Assume reactions.
                jsonRow.reactions_top_industry = value;
              } else {
                // Otherwise, assume comments.
                jsonRow.comments_top_industry = value;
              }

              // Early exit.
              return;
            }

            // Update.
            jsonRow[key] = value;
          });
        }
      }

      // Sheet two exists: YES.
      if (sheetTwo) {
        // Get sheet data.
        const sheetDataTwo = xl.utils.sheet_to_json(sheetTwo, { header: 1 });

        // Is array: YES.
        if (Array.isArray(sheetDataTwo)) {
          // Loop through.
          sheetDataTwo.forEach(([keyStart, keyEnd, value]) => {
            // Clean up.
            keyStart = String(keyStart || "")
              .trim()
              .toLowerCase()
              .replace(/\s+/g, "_");

            keyEnd = String(keyEnd || "")
              .trim()
              .toLowerCase()
              .replace(/\s+/g, "_")
              .replaceAll("-", "_to_")
              .replace(/\W/g, "")
              .replace("_employees", "");

            // Company size?
            if (keyStart === "company_size") {
              // Build key.
              let key = `employees_${keyEnd}`;

              // Suffix for large company.
              key = key.endsWith("10001") ? `${key}_or_more` : key;

              // Update.
              jsonRow[key] = value;
            }
          });
        }
      }

      // Add to list.
      jsonList.push(formatRow(jsonRow));
    } catch (error) {
      // Log.
      console.error(`Error processing file: ${filePath}`, error?.message);

      // Early exit.
      return;
    }
  });

  // Sort.
  jsonList.sort(sortByRecentDate);

  // Log example.
  console.log("============");
  console.log("Example row:");
  console.log("============");
  console.log("");
  console.log(
    // Format CLI output.
    JSON.stringify(jsonList[0], null, 2)
  );

  // Create new sheet & book.
  const newSheet = xl.utils.json_to_sheet(jsonList);
  const newCsv = xl.utils.sheet_to_csv(newSheet);
  const newExcel = xl.utils.book_new();

  // Add sheet.
  xl.utils.book_append_sheet(newExcel, newSheet, "Sheet1");

  // Output directory.
  const dirOutput = path.join(__dirname, "../output");

  // Path exists: NO.
  if (!fs.existsSync(dirOutput)) {
    // Create directory.
    fs.mkdirSync(dirOutput, { recursive: true });
  }

  // Create files.
  xl.writeFile(newExcel, fileOutputExcel);
  fs.writeFileSync(fileOutputCsv, newCsv, "utf8");

  // Log.
  console.log("");
  console.log("==================");
  console.log("New files created:");
  console.log("==================");
  console.log("");
  console.log(`• ${fileOutputCsv}`);
  console.log(`• ${fileOutputExcel}`);
  console.log("");
});
