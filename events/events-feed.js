/***********************
 * CONFIG
 ***********************/
const FEED_SHEET_NAME = "Feed";
const CSS_SHEET_NAMES = ["Settings", "CSS", "Styles"];
const ITEMS_PER_PAGE = 6;

/***********************
 * WEB APP ENTRY POINT
 ***********************/
function doGet() {
  const html = buildEventFeedHtml_();

  return HtmlService
    .createHtmlOutput(html)
    .setTitle("FDOB Event & Checklist Feed")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***********************
 * BUILD FULL HTML
 ***********************/
function buildEventFeedHtml_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = getFeedRows_(ss);
  const css = getCss_(ss);
  const filters = getFilterOptions_(rows);

  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">

  <style>
    ${css}

    .card:hover .button.secondary {
      background-color: #202124;
      color: white;
    }
  </style>
</head>

<body>
  <div class="container">
    <div class="event-container">
      <div class="event-filters">
        <div class="w-form">
          <form id="email-form" name="email-form" method="get" class="event-filter-form" aria-label="Event Filter Form">
            ${buildFilterGroup_("Region:", "region", filters.regions, "all")}
            ${buildFilterGroup_("Countries", "countries", filters.countries, "all")}
            ${buildFilterGroup_("Type", "type", filters.types, "all")}
            ${buildFilterGroup_("Month", "month", filters.months, "all")}
          </form>

          <div class="hide w-form-done" tabindex="-1" role="region" aria-label="Email Form success">
            <div>Thank you! Your submission has been received!</div>
          </div>

          <div class="hide w-form-fail" tabindex="-1" role="region" aria-label="Email Form failure">
            <div>Oops! Something went wrong while submitting the form.</div>
          </div>
        </div>
      </div>

      <div class="event-collection-wrapper">
        <div class="event-collection">
          <div class="empty" style="display: none;">
            <div>No Results</div>
          </div>

          ${rows.map(buildCard_).join("")}
        </div>

        <div class="event-pagination">
          <div class="button icon secondary disabled" style="pointer-events: none; opacity: 0.4;">
            <div class="tag-icon w-embed">
              ${ARROW_LEFT_SVG}
            </div>
          </div>

          <div class="button icon secondary">
            <div class="tag-icon w-embed">
              ${ARROW_RIGHT_SVG}
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  ${buildClientScript_()}
</body>
</html>
`;
}

/***********************
 * GET SHEET DATA
 ***********************/
function getFeedRows_(ss) {
  const sheet = ss.getSheetByName(FEED_SHEET_NAME) || ss.getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) return [];

  const headers = values[0].map(header => cleanHeader_(header));

  return values.slice(1)
    .map(row => {
      const item = {};

      headers.forEach((header, index) => {
        item[header] = row[index] || "";
      });

      return {
        date: item.date || "",
        month: item.month || "",
        region: item.region || "",
        type: item.type || "",
        country: item.country || "",
        title: item.title || "",
        description: item.description || "",
        link: item.link || ""
      };
    })
    .filter(item => {
      return (
        item.title ||
        item.description ||
        item.date ||
        item.month ||
        item.region ||
        item.type ||
        item.country
      );
    });
}

/***********************
 * GET CSS FROM SETTINGS TAB
 ***********************/
function getCss_(ss) {
  const cssSheet = CSS_SHEET_NAMES
    .map(name => ss.getSheetByName(name))
    .find(Boolean);

  if (!cssSheet) return "";

  const values = cssSheet.getDataRange().getDisplayValues();

  return values
    .flat()
    .filter(Boolean)
    .join("\n")
    .replace(/""/g, '"');
}

/***********************
 * BUILD FILTERS
 ***********************/
function getFilterOptions_(rows) {
  return {
    regions: uniqueValues_(rows.map(row => row.region)),
    // Split comma-separated countries so "IN, SEAS" becomes IN, SEAS in the filter.
    countries: uniqueValues_(
      rows.flatMap(row => splitMultiValue_(row.country))
    ),
    types: uniqueValues_(rows.map(row => row.type)),

    // Split comma-separated months so "June, July, August"
    // becomes separate filter options: June, July, August.
    months: sortMonths_(
      uniqueValues_(
        rows.flatMap(row => splitMultiValue_(row.month))
      )
    )
  };
}

function buildFilterGroup_(label, name, values, checkedValue) {
  const options = ["All", ...values];

  return `
    <div>
      <div class="text-size-xlarge text-weight-medium">${escapeHtml_(label)}</div>

      <div class="event-radio-filter">
        ${options.map((option, index) => {
          const value = option.toLowerCase() === "all" ? "all" : option;
          const id = `${name}-${slugify_(option)}-${index}`;
          const checked = value.toLowerCase() === checkedValue.toLowerCase();

          return `
            <label class="radio ${checked ? "checked" : ""} w-radio">
              <div class="w-form-formradioinput w-form-formradioinput--inputType-custom radio-button w-radio-input ${checked ? "w--redirected-checked" : ""}"></div>

              <input 
                type="radio" 
                data-name="${escapeHtml_(name)}" 
                id="${escapeHtml_(id)}" 
                name="${escapeHtml_(name)}" 
                style="opacity:0;position:absolute;z-index:-1" 
                ${checked ? "checked" : ""} 
                value="${escapeHtml_(value)}"
              >

              <span class="radio-button-label w-form-label" for="${escapeHtml_(id)}">${escapeHtml_(option)}</span>
            </label>
          `;
        }).join("")}
      </div>
    </div>
  `;
}

/***********************
 * BUILD CARDS
 ***********************/
function buildCard_(item) {
  const type = item.type || "";
  const typeClean = cleanValue_(type);
  const isEvent = typeClean === "event";
  const isChecklist = typeClean === "checklist";
  const link = normalizeLink_(item.link);
  const hasLink = link !== "";

  // Store separate month tags for filtering.
  // Example: "June, July, August" becomes "June|July|August"
  const monthTags = splitMultiValue_(item.month).join("|");

  // Same for countries: "IN, SEAS" -> "IN|SEAS" for OR-style filter matching.
  const countryTags = splitMultiValue_(item.country).join("|");

  const cardTag = hasLink ? "a" : "div";

  return `
    <${cardTag}
      data-date="${escapeHtml_(item.date)}"
      data-region="${escapeHtml_(item.region)}"
      data-country="${escapeHtml_(item.country)}"
      data-countries="${escapeHtml_(countryTags)}"
      data-type="${escapeHtml_(item.type)}"
      data-month="${escapeHtml_(item.month)}"
      data-months="${escapeHtml_(monthTags)}"
      ${hasLink ? `href="${escapeHtml_(link)}"` : ""}
      class="card w-inline-block"
      ${hasLink ? `target="_blank" rel="noopener noreferrer"` : ""}
    >
      <div>
        <div class="card-tag-wrapper">

          <div class="tag type-event" style="${isEvent ? "" : "display: none;"}">
            <div class="tag-icon w-embed">${EVENT_SVG}</div>
            <div>Event</div>
          </div>

          <div class="tag type-checklist" style="${isChecklist ? "" : "display: none;"}">
            <div class="tag-icon w-embed">${CHECKLIST_SVG}</div>
            <div>Checklist</div>
          </div>

          <div class="tag date" style="${item.date ? "" : "display: none;"}">
            <div class="tag-icon w-embed">${DATE_SVG}</div>
            <div>${escapeHtml_(item.date)}</div>
          </div>

          <div class="tag region" style="${item.region ? "" : "display: none;"}">
            <div class="tag-icon w-embed">${REGION_SVG}</div>
            <div>${escapeHtml_(item.region)}</div>
          </div>

        </div>

        <div class="card-text-wrapper">
          <div class="display-xs">${escapeHtml_(item.title)}</div>
          <div>${escapeHtml_(item.description)}</div>
        </div>
      </div>

      ${hasLink ? `
        <div class="button secondary">
          <div>Learn more</div>
        </div>
      ` : ""}
    </${cardTag}>
  `;
}

/***********************
 * FRONT-END FILTERING + PAGINATION
 ***********************/
function buildClientScript_() {
  return `
<script>
document.addEventListener("DOMContentLoaded", function () {
  const form = document.querySelector(".event-filter-form");
  const collection = document.querySelector(".event-collection");
  const cards = document.querySelectorAll(".event-collection .card");
  const emptyState = document.querySelector(".event-collection .empty");
  const pagination = document.querySelector(".event-pagination");
  const paginationButtons = pagination ? pagination.querySelectorAll(".button") : [];

  const prevButton = paginationButtons[0];
  const nextButton = paginationButtons[1];

  const itemsPerPage = ${ITEMS_PER_PAGE};
  let currentPage = 1;
  let filteredCards = [];

  if (!form || !collection) return;

  function clean(value) {
    return (value || "").trim().toLowerCase();
  }

  function getSelectedValue(name) {
    const checkedInput = form.querySelector('input[name="' + name + '"]:checked');
    return checkedInput ? clean(checkedInput.value) : "all";
  }

  function getMultiValues(value) {
    return String(value || "")
      .split("|")
      .map(function (item) {
        return clean(item);
      })
      .filter(Boolean);
  }

  function updateCheckedClasses() {
    const radioInputs = form.querySelectorAll('input[type="radio"]');

    radioInputs.forEach(function (input) {
      const radioLabel = input.closest(".radio");
      const fakeRadio = radioLabel ? radioLabel.querySelector(".radio-button") : null;

      if (!radioLabel) return;

      if (input.checked) {
        radioLabel.classList.add("checked");
        if (fakeRadio) fakeRadio.classList.add("w--redirected-checked");
      } else {
        radioLabel.classList.remove("checked");
        if (fakeRadio) fakeRadio.classList.remove("w--redirected-checked");
      }
    });
  }

  function getFilteredCards() {
    const selectedRegion = getSelectedValue("region");
    const selectedCountry = getSelectedValue("countries");
    const selectedType = getSelectedValue("type");
    const selectedMonth = getSelectedValue("month");

    return Array.from(cards).filter(function (card) {
      const cardRegion = clean(card.dataset.region);
      const cardType = clean(card.dataset.type);

      // data-countries: "IN|SEAS" -> ["in", "seas"] for multi-country rows
      const cardCountries = getMultiValues(card.dataset.countries);

      // Uses data-months: "June|July|August" -> ["june", "july", "august"]
      const cardMonths = getMultiValues(card.dataset.months);

      const regionMatch =
        selectedRegion === "all" ||
        cardRegion === selectedRegion;

      const countryMatch =
        selectedCountry === "all" ||
        cardCountries.includes(selectedCountry);

      const typeMatch =
        selectedType === "all" ||
        cardType === selectedType;

      const monthMatch =
        selectedMonth === "all" ||
        cardMonths.includes(selectedMonth);

      return regionMatch && countryMatch && typeMatch && monthMatch;
    });
  }

  function updatePaginationButtons(totalPages) {
    if (!pagination) return;

    if (totalPages <= 1) {
      pagination.style.display = "none";
      return;
    }

    pagination.style.display = "";

    if (prevButton) {
      const isDisabled = currentPage === 1;

      prevButton.classList.toggle("disabled", isDisabled);
      prevButton.style.pointerEvents = isDisabled ? "none" : "";
      prevButton.style.opacity = isDisabled ? "0.4" : "";
    }

    if (nextButton) {
      const isDisabled = currentPage === totalPages;

      nextButton.classList.toggle("disabled", isDisabled);
      nextButton.style.pointerEvents = isDisabled ? "none" : "";
      nextButton.style.opacity = isDisabled ? "0.4" : "";
    }
  }

  function renderCards() {
    const totalResults = filteredCards.length;
    const totalPages = Math.ceil(totalResults / itemsPerPage);

    cards.forEach(function (card) {
      card.style.display = "none";
    });

    if (emptyState) {
      emptyState.style.display = totalResults === 0 ? "" : "none";
    }

    if (totalResults === 0) {
      updatePaginationButtons(0);
      return;
    }

    if (currentPage > totalPages) {
      currentPage = totalPages;
    }

    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;

    filteredCards.slice(startIndex, endIndex).forEach(function (card) {
      card.style.display = "";
    });

    updatePaginationButtons(totalPages);
  }

  function filterCards() {
    filteredCards = getFilteredCards();
    currentPage = 1;
    renderCards();
  }

  function updateEverything() {
    updateCheckedClasses();
    filterCards();
  }

  form.addEventListener("change", updateEverything);

  form.addEventListener("submit", function (e) {
    e.preventDefault();
  });

  if (prevButton) {
    prevButton.addEventListener("click", function () {
      if (currentPage > 1) {
        currentPage--;
        renderCards();
      }
    });
  }

  if (nextButton) {
    nextButton.addEventListener("click", function () {
      const totalPages = Math.ceil(filteredCards.length / itemsPerPage);

      if (currentPage < totalPages) {
        currentPage++;
        renderCards();
      }
    });
  }

  updateEverything();
});
</script>
`;
}

/***********************
 * HELPERS
 ***********************/
function cleanHeader_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function cleanValue_(value) {
  return String(value || "").trim().toLowerCase();
}

function splitMultiValue_(value) {
  return String(value || "")
    .split(",")
    .map(item => item.trim())
    .filter(Boolean);
}

function uniqueValues_(values) {
  return [...new Set(
    values
      .map(value => String(value || "").trim())
      .filter(Boolean)
  )];
}

function sortMonths_(months) {
  const order = [
    "january",
    "february",
    "march",
    "april",
    "may",
    "june",
    "july",
    "august",
    "september",
    "october",
    "november",
    "december"
  ];

  return months.sort((a, b) => {
    const aIndex = order.indexOf(String(a).toLowerCase());
    const bIndex = order.indexOf(String(b).toLowerCase());

    if (aIndex === -1 && bIndex === -1) return a.localeCompare(b);
    if (aIndex === -1) return 1;
    if (bIndex === -1) return -1;

    return aIndex - bIndex;
  });
}

function slugify_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "");
}

function normalizeLink_(link) {
  const value = String(link || "").trim();

  if (!value) return "";
  if (value === "#") return "";

  if (
    value.startsWith("http://") ||
    value.startsWith("https://") ||
    value.startsWith("mailto:")
  ) {
    return value;
  }

  return "https://" + value;
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/***********************
 * SVG ICONS
 ***********************/
const EVENT_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none">
  <mask id="mask-event" style="mask-type:alpha" maskUnits="userSpaceOnUse" x="0" y="0" width="20" height="20">
    <rect width="20" height="20" fill="#D9D9D9"></rect>
  </mask>
  <g mask="url(#mask-event)">
    <path d="M12.5 10.3525L13.8142 11.6667V12.5H10.4167V16.6667L10 17.0833L9.58335 16.6667V12.5H6.18585V11.6667L7.50002 10.3525V4.16667H6.66669V3.33333H13.3334V4.16667H12.5V10.3525ZM7.37502 11.6667H12.625L11.6667 10.7083V4.16667H8.33335V10.7083L7.37502 11.6667Z" fill="white"></path>
  </g>
</svg>
`;

const CHECKLIST_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none">
  <mask id="mask-checklist" style="mask-type:alpha" maskUnits="userSpaceOnUse" x="0" y="0" width="20" height="20">
    <rect width="20" height="20" fill="#D9D9D9"></rect>
  </mask>
  <g mask="url(#mask-checklist)">
    <path d="M4.86542 14.9198L2.5 12.5546L3.07375 11.9808L4.84458 13.7517L8.38625 10.21L8.96 10.8046L4.86542 14.9198ZM4.86542 8.89418L2.5 6.52876L3.07375 5.95522L4.84458 7.72605L8.38625 4.18439L8.96 4.77876L4.86542 8.89418ZM10.8494 13.4296V12.5963H17.516V13.4296H10.8494ZM10.8494 7.40376V6.57043H17.516V7.40376H10.8494Z" fill="white"></path>
  </g>
</svg>
`;

const DATE_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none">
  <mask id="mask-date" style="mask-type:alpha" maskUnits="userSpaceOnUse" x="0" y="0" width="20" height="20">
    <rect width="20" height="20" fill="#D9D9D9"></rect>
  </mask>
  <g mask="url(#mask-date)">
    <path d="M11.1571 14.484C10.813 14.1399 10.641 13.7243 10.641 13.2371C10.641 12.75 10.813 12.3344 11.1571 11.9904C11.5011 11.6464 11.9166 11.4744 12.4037 11.4744C12.891 11.4744 13.3066 11.6464 13.6506 11.9904C13.9946 12.3344 14.1666 12.75 14.1666 13.2371C14.1666 13.7243 13.9946 14.1399 13.6506 14.484C13.3066 14.828 12.891 15 12.4037 15C11.9166 15 11.5011 14.828 11.1571 14.484ZM4.67956 17.5C4.29595 17.5 3.97567 17.3715 3.71873 17.1146C3.46179 16.8576 3.33331 16.5374 3.33331 16.1538V5.51292C3.33331 5.12931 3.46179 4.80903 3.71873 4.55208C3.97567 4.29514 4.29595 4.16667 4.67956 4.16667H6.15373V2.30771H7.05123V4.16667H13.0129V2.30771H13.8462V4.16667H15.3204C15.704 4.16667 16.0243 4.29514 16.2812 4.55208C16.5382 4.80903 16.6666 5.12931 16.6666 5.51292V16.1538C16.6666 16.5374 16.5382 16.8576 16.2812 17.1146C16.0243 17.3715 15.704 17.5 15.3204 17.5H4.67956ZM4.67956 16.6667H15.3204C15.4487 16.6667 15.5663 16.6133 15.6731 16.5065C15.7799 16.3997 15.8333 16.2821 15.8333 16.1538V8.84625H4.16665V16.1538C4.16665 16.2821 4.22005 16.3997 4.32685 16.5065C4.43366 16.6133 4.55123 16.6667 4.67956 16.6667ZM4.16665 8.01271H15.8333V5.51292C15.8333 5.38458 15.7799 5.26701 15.6731 5.16021C15.5663 5.0534 15.4487 5 15.3204 5H4.67956C4.55123 5 4.43366 5.0534 4.32685 5.16021C4.22005 5.26701 4.16665 5.38458 4.16665 5.51292V8.01271Z" fill="white"></path>
  </g>
</svg>
`;

const REGION_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20" fill="none">
  <mask id="mask0_4559_3321" style="mask-type:alpha" maskUnits="userSpaceOnUse" x="0" y="0" width="20" height="20">
    <rect width="20" height="20" fill="#D9D9D9"/>
  </mask>
  <g mask="url(#mask0_4559_3321)">
    <path d="M7.08333 16.9079C6.17097 16.5131 5.37639 15.9772 4.69958 15.3004C4.02278 14.6236 3.48694 13.829 3.09208 12.9167C2.69736 12.0043 2.5 11.0312 2.5 9.9975C2.5 8.96389 2.69736 7.99167 3.09208 7.08083C3.48694 6.17014 4.02278 5.37639 4.69958 4.69958C5.37639 4.02278 6.17097 3.48694 7.08333 3.09208C7.99569 2.69736 8.96875 2.5 10.0025 2.5C11.0361 2.5 12.0083 2.69736 12.9192 3.09208C13.8299 3.48694 14.6236 4.02278 15.3004 4.69958C15.9772 5.37639 16.5131 6.17014 16.9079 7.08083C17.3026 7.99167 17.5 8.96389 17.5 9.9975C17.5 11.0312 17.3026 12.0043 16.9079 12.9167C16.5131 13.829 15.9772 14.6236 15.3004 15.3004C14.6236 15.9772 13.8299 16.5131 12.9192 16.9079C12.0083 17.3026 11.0361 17.5 10.0025 17.5C8.96875 17.5 7.99569 17.3026 7.08333 16.9079ZM10 16.6731C10.4893 16.0449 10.8926 15.4279 11.21 14.8221C11.5272 14.2164 11.7852 13.5385 11.984 12.7885H8.01604C8.23618 13.5812 8.49951 14.2804 8.80604 14.8862C9.11271 15.4919 9.51069 16.0876 10 16.6731ZM8.93917 16.5481C8.55028 16.0898 8.19528 15.5233 7.87417 14.8485C7.55319 14.1738 7.31417 13.4872 7.15708 12.7885H3.96146C4.4391 13.8248 5.11618 14.6747 5.99271 15.3381C6.86937 16.0016 7.85153 16.4049 8.93917 16.5481ZM11.0608 16.5481C12.1485 16.4049 13.1306 16.0016 14.0073 15.3381C14.8838 14.6747 15.5609 13.8248 16.0385 12.7885H12.8429C12.6325 13.4978 12.3667 14.1899 12.0456 14.8646C11.7247 15.5393 11.3964 16.1005 11.0608 16.5481ZM3.62188 11.9552H6.98396C6.9209 11.6133 6.87632 11.2801 6.85021 10.9558C6.82396 10.6317 6.81083 10.3131 6.81083 10C6.81083 9.68694 6.82396 9.36833 6.85021 9.04417C6.87632 8.71986 6.9209 8.38674 6.98396 8.04479H3.62188C3.53104 8.33326 3.46021 8.64764 3.40937 8.98792C3.35868 9.32819 3.33333 9.66556 3.33333 10C3.33333 10.3344 3.35868 10.6718 3.40937 11.0121C3.46021 11.3524 3.53104 11.6667 3.62188 11.9552ZM7.81729 11.9552H12.1827C12.2458 11.6133 12.2903 11.2855 12.3165 10.9719C12.3427 10.6584 12.3558 10.3344 12.3558 10C12.3558 9.66556 12.3427 9.3416 12.3165 9.02812C12.2903 8.71451 12.2458 8.38674 12.1827 8.04479H7.81729C7.75424 8.38674 7.70965 8.71451 7.68354 9.02812C7.65729 9.3416 7.64417 9.66556 7.64417 10C7.64417 10.3344 7.65729 10.6584 7.68354 10.9719C7.70965 11.2855 7.75424 11.6133 7.81729 11.9552ZM13.016 11.9552H16.3781C16.469 11.6667 16.5398 11.3524 16.5906 11.0121C16.6413 10.6718 16.6667 10.3344 16.6667 10C16.6667 9.66556 16.6413 9.32819 16.5906 8.98792C16.5398 8.64764 16.469 8.33326 16.3781 8.04479H13.016C13.0791 8.38674 13.1237 8.71986 13.1498 9.04417C13.176 9.36833 13.1892 9.68694 13.1892 10C13.1892 10.3131 13.176 10.6317 13.1498 10.9558C13.1237 11.2801 13.0791 11.6133 13.016 11.9552ZM12.8429 7.21146H16.0385C15.5502 6.15382 14.8811 5.30396 14.0312 4.66188C13.1814 4.01979 12.1912 3.61111 11.0608 3.43583C11.4497 3.94764 11.7994 4.53285 12.1098 5.19146C12.4202 5.85021 12.6646 6.52354 12.8429 7.21146ZM8.01604 7.21146H11.984C11.7638 6.42951 11.4924 5.72229 11.1698 5.08979C10.8472 4.45729 10.4572 3.86965 10 3.32688C9.54278 3.86965 9.15285 4.45729 8.83021 5.08979C8.50757 5.72229 8.23618 6.42951 8.01604 7.21146ZM3.96146 7.21146H7.15708C7.33542 6.52354 7.57979 5.85021 7.89021 5.19146C8.20062 4.53285 8.55028 3.94764 8.93917 3.43583C7.79806 3.61111 6.80528 4.02243 5.96083 4.66979C5.11625 5.31729 4.44979 6.16451 3.96146 7.21146Z" fill="white"/>
  </g>
</svg>
`;

const ARROW_LEFT_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none">
  <path d="M19 12H5M12 5L5 12L12 19" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"></path>
</svg>
`;

const ARROW_RIGHT_SVG = `
<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none">
  <path d="M5 12H19M12 19L19 12L12 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"></path>
</svg>
`;