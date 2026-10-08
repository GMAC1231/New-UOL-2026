/* =========================================================
   UOL TENTATIVE DASHBOARD — APP.JS

   WORKBOOK:
   data/UOL_Separated_All_Activity_Data_Cleaned.xlsx

   CASCADING FILTERS:
   Faculty
      -> Department
      -> Activity Type
      -> Reporting Month
      -> Organization
   ========================================================= */

let allRows = [];
let filteredRows = [];

let allIpRows = [];
let ipWorkbookLoaded = false;

const charts = {};

const $ = id =>
  document.getElementById(id);


/* =========================================================
   HELPERS
   ========================================================= */

function clean(value) {

  if (
    value === undefined ||
    value === null
  ) {
    return "";
  }

  const text =
    String(value).trim();

  if (
    !text ||
    [
      "undefined",
      "null",
      "nan"
    ].includes(
      text.toLowerCase()
    )
  ) {
    return "";
  }

  return text;
}


function norm(value) {

  return clean(value)
    .toLowerCase()
    .replace(
      /[^a-z0-9]+/g,
      " "
    )
    .trim();
}


function normalizeDisplayText(value) {

  let text = clean(value);

  if (!text) {
    return "";
  }

  // Fix OCR / extracted text where letters are split with spaces,
  // e.g. "H e a l t h" -> "Health".
  text = text.replace(/\s*([,.;:!?()\/])\s*/g, "$1 ");
  text = text.replace(/\s*-\s*/g, "-");
  text = text.replace(/\s{2,}/g, " ").trim();

  const collapseSingleLetterWords = input =>
    input.replace(/\b(?:[A-Za-z]\s+){2,}[A-Za-z]\b/g, match =>
      match.replace(/\s+/g, "")
    );

  let prev = "";
  while (text !== prev) {
    prev = text;
    text = collapseSingleLetterWords(text);
  }

  // Tidy punctuation spacing after collapsing.
  text = text
    .replace(/\s+([,.;:!?])/g, "$1")
    .replace(/([,.;:!?])(\S)/g, "$1 $2")
    .replace(/\s{2,}/g, " ")
    .trim();

  return text;
}


function val(
  row,
  names
) {

  const keys =
    Object.keys(row);

  for (
    const name
    of names
  ) {

    const key =
      keys.find(
        k =>
          norm(k) ===
          norm(name)
      );

    if (
      key !== undefined &&
      clean(row[key])
    ) {

      return clean(
        row[key]
      );
    }
  }

  return "";
}


function num(value) {

  const match =
    clean(value)
      .replace(
        /,/g,
        ""
      )
      .match(
        /-?\d+(?:\.\d+)?/
      );

  return match
    ? Number(match[0])
    : 0;
}


function yes(value) {

  return [
    "yes",
    "y",
    "true",
    "1",
    "scheduled"
  ].includes(
    clean(value)
      .toLowerCase()
  );
}


function esc(value) {

  return String(
    value ?? ""
  )
    .replace(
      /&/g,
      "&amp;"
    )
    .replace(
      /</g,
      "&lt;"
    )
    .replace(
      />/g,
      "&gt;"
    )
    .replace(
      /"/g,
      "&quot;"
    );
}


function isInvalidLabel(value) {

  const text =
    norm(value);

  return (
    !text ||
    text === "not available" ||
    text === "unclassified" ||
    text === "undefined" ||
    text === "null" ||
    text === "n a" ||
    text === "na"
  );
}


/* =========================================================
   ACTIVITY TYPE STANDARDIZATION
   ========================================================= */

function standardizeActivityType(value) {

  const type =
    norm(value);

  if (
    type === "glit" ||
    type.includes(
      "guest lecture"
    )
  ) {
    return "Guest Lecture / GLIT";
  }

  if (
    type === "iv" ||
    type.includes(
      "industrial visit"
    )
  ) {
    return "Industrial Visit / IV";
  }

  if (
    type === "planned" ||
    type.includes(
      "planned activity"
    )
  ) {
    return "Planned";
  }

  if (
    type === "ip" ||
    type === "internship placement" ||
    type.includes(
      "internship"
    ) ||
    type.includes(
      "intership"
    ) ||
    type.includes(
      "placement"
    )
  ) {
    return "Internship / Placement";
  }

  if (
    type === "mou" ||
    type.includes(
      "mou signing"
    ) ||
    type.includes(
      "memorandum of understanding"
    )
  ) {
    return "MoU / MoU Signing";
  }

  if (
    type.includes(
      "iab"
    ) ||
    type.includes(
      "industry consultation"
    )
  ) {
    return "IAB / Industry Consultation";
  }

  if (
    type.includes(
      "research collaboration"
    )
  ) {
    return "Research Collaboration";
  }

  if (
    type.includes(
      "community engagement"
    )
  ) {
    return "Community Engagement";
  }

  if (
    type.includes(
      "alumni"
    )
  ) {
    return "Alumni Talk / Mentoring";
  }

  if (
    type.includes(
      "curriculum feedback"
    )
  ) {
    return "Curriculum Feedback";
  }

  return clean(value);
}


/* =========================================================
   PARSE MAIN ACTIVITY DATA
   ========================================================= */

function parseActivities(rows) {

  const parsed =
    rows.map(
      row => {

        const rawType =
          val(
            row,
            [
              "Separated Activity Type",
              "Activity Type"
            ]
          );

        return {

          raw:
            row,

          ref:
            val(
              row,
              [
                "Ref #",
                "Ref",
                "Reference"
              ]
            ),

          month:
            val(
              row,
              [
                "Reporting Month",
                "Month",
                "Submission Date"
              ]
            ),

          faculty:
            val(
              row,
              [
                "Faculty"
              ]
            ),

          department:
            val(
              row,
              [
                "Department"
              ]
            ),

          person:
            val(
              row,
              [
                "Focal Person",
                "Submitted By",
                "Person",
                "Name"
              ]
            ),

          designation:
            val(
              row,
              [
                "Designation"
              ]
            ),

          originalType:
            rawType,

          type:
            standardizeActivityType(
              rawType
            ),

          events:
            num(
              val(
                row,
                [
                  "How Many",
                  "Planned Quantity",
                  "Total Activities"
                ]
              )
            ),

          organization:
            val(
              row,
              [
                "Organization / Activity Title",
                "Organization / Activity",
                "Name Of Organization Planned with",
                "Name of Organization Planned with",
                "Planned Partner / Organization",
                "Organization"
              ]
            ),

          planned:
            val(
              row,
              [
                "Planned",
                "Planned?"
              ]
            ),

          scheduled:
            val(
              row,
              [
                "Scheduled",
                "Scheduled?"
              ]
            ),

          remarks:
            val(
              row,
              [
                "Remarks / Result",
                "Remarks/Result",
                "Actual Outcome",
                "Remarks"
              ]
            )
        };
      }
    );

  return parsed.filter(
    row => (
      !isInvalidLabel(
        row.faculty
      ) &&
      !isInvalidLabel(
        row.department
      ) &&
      !isInvalidLabel(
        row.type
      ) &&
      norm(row.type) !== "total activities"
    )
  );
}


/* =========================================================
   INTERNSHIP / PLACEMENT BREAKDOWN
   ========================================================= */

function buildIpBreakdownFromActivities() {

  const grouped =
    {};

  allRows.forEach(
    row => {

      if (
        row.type !==
        "Internship / Placement"
      ) {
        return;
      }

      if (
        row.events <= 0
      ) {
        return;
      }

      let sourceType =
        "Internship/Placement";

      if (
        norm(
          row.originalType
        ) === "ip"
      ) {
        sourceType =
          "IP";
      }

      const semester =
        semesterFromMonth(
          row.month
        );

      const key =
        [
          row.faculty,
          row.department,
          sourceType,
          semester
        ].join(
          "||"
        );

      if (
        !grouped[key]
      ) {

        grouped[key] = {

          faculty:
            row.faculty,

          department:
            row.department,

          sourceType:
            sourceType,

          semester:
            semester,

          students:
            0
        };
      }

      grouped[key]
        .students +=
        row.events;
    }
  );

  allIpRows =
    Object.values(
      grouped
    );

  ipWorkbookLoaded =
    true;
}


/* =========================================================
   SELECT HELPERS
   ========================================================= */

function unique(values) {

  return [
    ...new Set(
      values
        .map(
          clean
        )
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    )
  ].sort(
    (
      a,
      b
    ) =>
      a.localeCompare(
        b
      )
  );
}


function setSelectOptions(
  id,
  values,
  allLabel
) {

  const select =
    $(id);

  if (
    !select
  ) {
    return;
  }

  const previous =
    select.value;

  const items =
    unique(
      values
    );

  select.innerHTML =
    `<option value="All">${allLabel}</option>`;

  items.forEach(
    item => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        item;

      option.textContent =
        item;

      select.appendChild(
        option
      );
    }
  );

  if (
    previous !== "All" &&
    items.includes(
      previous
    )
  ) {

    select.value =
      previous;

  } else {

    select.value =
      "All";
  }
}


/* =========================================================
   SEMESTER-FIRST CASCADING FILTERS

   Semester
      -> Faculty
      -> Department
      -> Activity Type
      -> Organization
   ========================================================= */

function currentSemester() {
  return $("monthFilter")?.value || "All";
}

function rowsForSemester(rows = allRows) {
  const semester = currentSemester();

  if (semester === "All") {
    return [...rows];
  }

  return rows.filter(
    row => semesterFromMonth(row.month) === semester
  );
}


/* =========================================================
   FACULTY FILTER
   Depends on Semester
   ========================================================= */

function updateFacultyFilter() {

  const rows = rowsForSemester(allRows);

  setSelectOptions(
    "facultyFilter",
    rows.map(row => row.faculty),
    "All Faculties"
  );
}


/* =========================================================
   DEPARTMENT FILTER
   Depends on Semester + Faculty
   ========================================================= */

function updateDepartmentFilter() {

  const faculty =
    $("facultyFilter")?.value || "All";

  let rows = rowsForSemester(allRows);

  if (faculty !== "All") {
    rows = rows.filter(
      row => row.faculty === faculty
    );
  }

  setSelectOptions(
    "departmentFilter",
    rows.map(row => row.department),
    "All Departments"
  );
}


/* =========================================================
   ACTIVITY FILTER
   Depends on Semester + Faculty + Department
   ========================================================= */

function updateActivityFilter() {

  const faculty =
    $("facultyFilter")?.value || "All";

  const department =
    $("departmentFilter")?.value || "All";

  let rows = rowsForSemester(allRows);

  if (faculty !== "All") {
    rows = rows.filter(
      row => row.faculty === faculty
    );
  }

  if (department !== "All") {
    rows = rows.filter(
      row => row.department === department
    );
  }

  setSelectOptions(
    "activityFilter",
    rows
      .map(row => row.type)
      .filter(type => norm(type) !== "total activities"),
    "All Activity Types"
  );
}


/* =========================================================
   SEMESTER FILTER
   ========================================================= */

function semesterFromMonth(value) {
  const text = clean(value);
  if (!text) return "";

  const lower = text.toLowerCase();
  const yearMatch = text.match(/\b(20\d{2})\b/);
  const explicitYear = yearMatch ? yearMatch[1] : "";

  if (lower.includes("spring")) {
    return explicitYear ? `Spring Semester ${explicitYear}` : "Spring Semester";
  }

  if (lower.includes("fall") || lower.includes("autumn")) {
    return explicitYear ? `Fall Semester ${explicitYear}` : "Fall Semester";
  }

  const monthNames = {
    january: 1, jan: 1,
    february: 2, feb: 2,
    march: 3, mar: 3,
    april: 4, apr: 4,
    may: 5,
    june: 6, jun: 6,
    july: 7, jul: 7,
    august: 8, aug: 8,
    september: 9, sept: 9, sep: 9,
    october: 10, oct: 10,
    november: 11, nov: 11,
    december: 12, dec: 12
  };

  let month = 0;
  let year = explicitYear;

  const iso = text.match(/\b(20\d{2})[-\/](\d{1,2})(?:[-\/]\d{1,2})?\b/);
  if (iso) {
    year = iso[1];
    month = Number(iso[2]);
  }

  if (!month) {
    const m = lower.match(/\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b/);
    if (m) month = monthNames[m[1]] || 0;
  }

  if (!month) {
    const date = new Date(text);
    if (!Number.isNaN(date.getTime())) {
      month = date.getMonth() + 1;
      year = String(date.getFullYear());
    }
  }

  if (!month) return text;

  const semester = month <= 6 ? "Spring Semester" : "Fall Semester";
  return year ? `${semester} ${year}` : semester;
}

function updateMonthFilter() {

  setSelectOptions(
    "monthFilter",
    allRows
      .map(row => semesterFromMonth(row.month))
      .filter(Boolean),
    "All Semesters"
  );
}


/* =========================================================
   ORGANIZATION FILTER
   Depends on Semester + Faculty + Department + Activity Type
   ========================================================= */

function updateOrganizationFilter() {

  const faculty =
    $("facultyFilter")?.value || "All";

  const department =
    $("departmentFilter")?.value || "All";

  const activity =
    $("activityFilter")?.value || "All";

  let rows = rowsForSemester(allRows);

  if (faculty !== "All") {
    rows = rows.filter(
      row => row.faculty === faculty
    );
  }

  if (department !== "All") {
    rows = rows.filter(
      row => row.department === department
    );
  }

  if (activity !== "All") {
    rows = rows.filter(
      row => row.type === activity
    );
  }

  setSelectOptions(
    "organizationFilter",
    rows
      .map(row => row.organization)
      .filter(organization => !isInvalidLabel(organization)),
    "All Organizations"
  );
}


/* =========================================================
   INITIAL FILTER LOAD
   ========================================================= */

function populateFilters() {

  updateMonthFilter();

  if ($("monthFilter")) {
    $("monthFilter").value = "All";
  }

  updateFacultyFilter();

  if ($("facultyFilter")) {
    $("facultyFilter").value = "All";
  }

  updateDepartmentFilter();
  updateActivityFilter();
  updateOrganizationFilter();
}


/* =========================================================
   APPLY FILTERS
   ========================================================= */

function applyFilters() {

  const faculty =
    $("facultyFilter")?.value || "All";

  const department =
    $("departmentFilter")?.value || "All";

  const activity =
    $("activityFilter")?.value || "All";

  const month =
    $("monthFilter")?.value || "All";

  const organization =
    $("organizationFilter")?.value || "All";

  const search =
    ($("searchFilter")?.value || "")
      .toLowerCase()
      .trim();

  filteredRows = allRows.filter(row => {

    const facultyMatch =
      faculty === "All" || row.faculty === faculty;

    const departmentMatch =
      department === "All" || row.department === department;

    const activityMatch =
      activity === "All" || row.type === activity;

    const monthMatch =
      month === "All" || semesterFromMonth(row.month) === month;

    const organizationMatch =
      organization === "All" || row.organization === organization;

    const searchMatch =
      !search || [
        row.faculty,
        row.department,
        row.type,
        row.originalType,
        row.person,
        row.designation,
        row.organization,
        row.remarks,
        row.month,
        semesterFromMonth(row.month)
      ].some(value =>
        String(value || "")
          .toLowerCase()
          .includes(search)
      );

    return (
      facultyMatch &&
      departmentMatch &&
      activityMatch &&
      monthMatch &&
      organizationMatch &&
      searchMatch
    );
  });

  render();
}


/* =========================================================
   FILTER HANDLERS
   ========================================================= */

function handleMonthChange() {

  /*
    Strict cascading order:
      Semester -> Faculty -> Department -> Activity Type -> Organization

    When an upstream filter changes, all downstream selections are reset
    so stale values cannot remain active.
  */

  updateFacultyFilter();
  if ($("facultyFilter")) $("facultyFilter").value = "All";

  updateDepartmentFilter();
  if ($("departmentFilter")) $("departmentFilter").value = "All";

  updateActivityFilter();
  if ($("activityFilter")) $("activityFilter").value = "All";

  updateOrganizationFilter();
  if ($("organizationFilter")) $("organizationFilter").value = "All";

  applyFilters();
}

function handleFacultyChange() {

  updateDepartmentFilter();
  if ($("departmentFilter")) $("departmentFilter").value = "All";

  updateActivityFilter();
  if ($("activityFilter")) $("activityFilter").value = "All";

  updateOrganizationFilter();
  if ($("organizationFilter")) $("organizationFilter").value = "All";

  applyFilters();
}

function handleDepartmentChange() {

  updateActivityFilter();
  if ($("activityFilter")) $("activityFilter").value = "All";

  updateOrganizationFilter();
  if ($("organizationFilter")) $("organizationFilter").value = "All";

  applyFilters();
}

function handleActivityChange() {

  updateOrganizationFilter();
  if ($("organizationFilter")) $("organizationFilter").value = "All";

  applyFilters();
}

function handleOrganizationChange() {
  applyFilters();
}


/* =========================================================
   RESET
   ========================================================= */

function resetFilters() {

  if ($("monthFilter")) {
    $("monthFilter").value = "All";
  }

  updateMonthFilter();
  updateFacultyFilter();

  if ($("facultyFilter")) {
    $("facultyFilter").value = "All";
  }

  updateDepartmentFilter();

  if ($("departmentFilter")) {
    $("departmentFilter").value = "All";
  }

  updateActivityFilter();

  if ($("activityFilter")) {
    $("activityFilter").value = "All";
  }

  updateOrganizationFilter();

  if ($("organizationFilter")) {
    $("organizationFilter").value = "All";
  }

  if ($("searchFilter")) {
    $("searchFilter").value = "";
  }

  filteredRows = [...allRows];
  render();
}

/* =========================================================
   INTERNSHIP / PLACEMENT
   ========================================================= */

function getFilteredIpRows() {

  if (
    !ipWorkbookLoaded
  ) {
    return [];
  }

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const department =
    $("departmentFilter")
      ?.value ||
    "All";

  const semester =
    $("monthFilter")
      ?.value ||
    "All";

  return allIpRows.filter(
    row => {

      const facultyMatch =
        (
          faculty === "All" ||
          row.faculty ===
          faculty
        );

      const departmentMatch =
        (
          department === "All" ||
          row.department ===
          department
        );

      const semesterMatch =
        (
          semester === "All" ||
          row.semester ===
          semester
        );

      return (
        facultyMatch &&
        departmentMatch &&
        semesterMatch
      );
    }
  );
}


function summarizeIpRows(rows) {

  const byType =
    {};

  let total =
    0;

  rows.forEach(
    row => {

      if (
        row.students <= 0
      ) {
        return;
      }

      total +=
        row.students;

      const key =
        clean(
          row.sourceType
        );

      byType[key] =
        (
          byType[key] ||
          0
        ) +
        row.students;
    }
  );

  const departments =
    new Set(
      rows
        .map(
          row =>
            row.department
        )
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    );

  return {

    total:
      total,

    byType:
      byType,

    rowCount:
      rows.length,

    departmentCount:
      departments.size
  };
}


function setImpactRowVisible(
  id,
  visible
) {

  const element =
    $(id);

  if (
    !element
  ) {
    return;
  }

  const row =
    element.closest(
      ".impact-row"
    );

  if (
    row
  ) {

    row.style.display =
      visible
        ? ""
        : "none";
  }
}


function setImpactLabel(
  id,
  text
) {

  const element =
    $(id);

  if (
    !element
  ) {
    return;
  }

  const row =
    element.closest(
      ".impact-row"
    );

  const label =
    row
      ?.querySelector(
        "span"
      );

  if (
    label
  ) {

    label.textContent =
      text;
  }
}


function renderStudentSummary() {

  const rows =
    getFilteredIpRows();

  const summary =
    summarizeIpRows(
      rows
    );

  [
    "kInternshipStudents",
    "sideStudents",
    "impactStudents",
    "impactCombined"
  ].forEach(
    id => {

      if (
        $(id)
      ) {

        $(id)
          .textContent =
          summary.total
            .toLocaleString();
      }
    }
  );

  if (
    $("impactIpRecords")
  ) {

    $("impactIpRecords")
      .textContent =
      summary.rowCount
        .toLocaleString();
  }

  if (
    $("impactIpDepartments")
  ) {

    $("impactIpDepartments")
      .textContent =
      summary.departmentCount
        .toLocaleString();
  }

  const slots =
    [
      "impactInternship",
      "impactPlacement",
      "impactPlanned"
    ];

  slots.forEach(
    id =>
      setImpactRowVisible(
        id,
        false
      )
  );

  const typeEntries =
    Object.entries(
      summary.byType
    )
      .filter(
        (
          [
            ,
            value
          ]
        ) =>
          value > 0
      )
      .sort(
        (
          a,
          b
        ) =>
          b[1] -
          a[1]
      );

  typeEntries
    .slice(
      0,
      slots.length
    )
    .forEach(
      (
        [
          type,
          value
        ],
        index
      ) => {

        const id =
          slots[
            index
          ];

        setImpactLabel(
          id,

          type === "IP"
            ? "IP Students"
            : `${type} Students`
        );

        if (
          $(id)
        ) {

          $(id)
            .textContent =
            value
              .toLocaleString();
        }

        setImpactRowVisible(
          id,
          true
        );
      }
    );
}


/* =========================================================
   CHART DATA
   ========================================================= */

function rowAmount(row) {

  return row.events > 0
    ? row.events
    : 1;
}


function group(key) {

  const result =
    {};

  filteredRows.forEach(
    row => {

      const label =
        clean(
          row[key]
        );

      if (
        isInvalidLabel(
          label
        )
      ) {
        return;
      }

      result[label] =
        (
          result[label] ||
          0
        ) +
        rowAmount(
          row
        );
    }
  );

  return Object.fromEntries(
    Object.entries(
      result
    )
      .sort(
        (
          a,
          b
        ) =>
          b[1] -
          a[1]
      )
  );
}


function destroyChart(id) {

  if (
    charts[id]
  ) {

    charts[id]
      .destroy();

    delete charts[id];
  }
}


function wrapLabel(
  text,
  max = 32
) {

  const words =
    String(text)
      .split(
        /\s+/
      );

  const lines =
    [];

  let line =
    "";

  words.forEach(
    word => {

      const next =
        line
          ? `${line} ${word}`
          : word;

      if (
        next.length >
          max &&
        line
      ) {

        lines.push(
          line
        );

        line =
          word;

      } else {

        line =
          next;
      }
    }
  );

  if (
    line
  ) {

    lines.push(
      line
    );
  }

  return lines;
}


/* =========================================================
   BAR CHART
   ========================================================= */

function barChart(
  id,
  data,
  key,
  wrapper
) {

  const canvas =
    $(id);

  if (
    !canvas
  ) {
    return;
  }

  destroyChart(
    id
  );

  const labels =
    Object.keys(
      data
    );

  const values =
    Object.values(
      data
    );

  const chartTotal =
    values.reduce(
      (sum, value) => sum + Number(value || 0),
      0
    );

  const chartPercent = value =>
    chartTotal > 0
      ? ((Number(value || 0) / chartTotal) * 100).toFixed(1) + "%"
      : "0.0%";

  if (
    wrapper &&
    $(wrapper)
  ) {

    $(wrapper)
      .style
      .height =
      Math.max(
        280,
        labels.length * 34
      ) +
      "px";
  }

  const visibleBarValuePlugin = {
    id: `visibleBarValue_${id}`,
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      ctx.save();
      ctx.font = "600 10px Inter, Segoe UI, Arial, sans-serif";
      ctx.fillStyle = "#334155";
      ctx.textAlign = "left";
      ctx.textBaseline = "middle";

      meta.data.forEach((bar, index) => {
        const value = Number(values[index] || 0);
        const label = `${value.toLocaleString()} (${chartPercent(value)})`;
        const x = Math.min(bar.x + 7, chart.chartArea.right + 68);
        ctx.fillText(label, x, bar.y);
      });
      ctx.restore();
    }
  };

  charts[id] =
    new Chart(
      canvas,
      {

        type:
          "bar",

        plugins: [visibleBarValuePlugin],

        data: {

          labels:
            labels,

          datasets: [

            {

              data:
                values,

              backgroundColor:

                labels.map(
                  (
                    _,
                    index
                  ) =>

                    [
                      "#0b6b3a",
                      "#1556c0",
                      "#7f2db4",
                      "#ef7d17",
                      "#0b7f90",
                      "#14206d"
                    ][
                      index %
                      6
                    ]
                ),

              borderRadius:
                5,

              maxBarThickness:
                19
            }
          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          layout: {
            padding: {
              right: 92
            }
          },

          indexAxis:
            "y",

          onClick:
            (
              event,
              elements
            ) => {

              if (
                !elements.length
              ) {
                return;
              }

              const index =
                elements[0]
                  .index;

              openPopup(
                key,
                labels[index]
              );
            },

          plugins: {

            legend: {

              display:
                false
            },

            tooltip: {
              callbacks: {
                label: context => {
                  const value = Number(context.raw || 0);
                  return `${value.toLocaleString()} (${chartPercent(value)})`;
                }
              }
            }
          },

          scales: {

            x: {

              beginAtZero:
                true,

              ticks: {

                precision:
                  0
              }
            },

            y: {

              ticks: {

                autoSkip:
                  false,

                callback:
                  function(
                    value
                  ) {

                    const label =
                      this
                        .getLabelForValue(
                          value
                        );

                    const index =
                      labels.indexOf(
                        label
                      );

                    const pct =
                      index >= 0
                        ? chartPercent(values[index])
                        : "0.0%";

                    return wrapLabel(
                      `${label} (${pct})`,
                      30
                    );
                  },

                font: {

                  size:
                    9
                }
              },

              grid: {

                display:
                  false
              }
            }
          }
        }
      }
    );
}


/* =========================================================
   ACTIVITY DISTRIBUTION
   ========================================================= */

function activityChart() {

  const canvas =
    $("activityChart");

  if (
    !canvas
  ) {
    return;
  }

  destroyChart(
    "activityChart"
  );

  const grouped =
    group(
      "type"
    );

  const entries =
    Object.entries(
      grouped
    )
      .filter(
        (
          [
            label,
            value
          ]
        ) =>
          value > 0 &&
          !isInvalidLabel(
            label
          )
      )
      .sort(
        (
          a,
          b
        ) =>
          b[1] -
          a[1]
      );

  const labels =
    entries.map(
      (
        [
          label
        ]
      ) =>
        label
    );

  const values =
    entries.map(
      (
        [
          ,
          value
        ]
      ) =>
        value
    );

  const activityTotal =
    values.reduce(
      (sum, value) => sum + Number(value || 0),
      0
    );

  const activityPercent = value =>
    activityTotal > 0
      ? ((Number(value || 0) / activityTotal) * 100).toFixed(1) + "%"
      : "0.0%";

  const palette =
    [
      "#0b6b3a",
      "#1556c0",
      "#7f2db4",
      "#ef7d17",
      "#0b7f90",
      "#14206d",
      "#e04b5a",
      "#d4a017",
      "#166534",
      "#2563eb",
      "#9333ea",
      "#ea580c",
      "#0e7490",
      "#4338ca",
      "#be123c",
      "#4f46e5"
    ];

  const colors =
    labels.map(
      (
        _,
        index
      ) =>
        palette[
          index %
          palette.length
        ]
    );

  const visibleActivityPercentPlugin = {
    id: "visibleActivityPercent",
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      ctx.save();
      ctx.font = "700 10px Inter, Segoe UI, Arial, sans-serif";
      ctx.fillStyle = "#ffffff";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      meta.data.forEach((arc, index) => {
        const value = Number(values[index] || 0);
        const numericPct = activityTotal > 0 ? (value / activityTotal) * 100 : 0;
        if (numericPct < 3) return;
        const pos = arc.tooltipPosition();
        ctx.fillText(`${numericPct.toFixed(1)}%`, pos.x, pos.y);
      });
      ctx.restore();
    }
  };

  charts.activityChart =
    new Chart(
      canvas,
      {

        type:
          "doughnut",

        plugins: [visibleActivityPercentPlugin],

        data: {

          labels:
            labels,

          datasets: [

            {

              data:
                values,

              backgroundColor:
                colors,

              borderColor:
                "#ffffff",

              borderWidth:
                2,

              hoverOffset:
                5
            }
          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          cutout:
            "60%",

          onClick:
            (
              event,
              elements
            ) => {

              if (
                !elements.length
              ) {
                return;
              }

              openPopup(
                "type",
                labels[
                  elements[0]
                    .index
                ]
              );
            },

          plugins: {

            legend: {

              display:
                false
            },

            tooltip: {
              callbacks: {
                label: context => {
                  const value = Number(context.raw || 0);
                  const label = context.label || "";
                  return `${label}: ${value.toLocaleString()} (${activityPercent(value)})`;
                }
              }
            }
          }
        }
      }
    );

  renderActivityLegend(
    labels,
    values,
    colors
  );
}


function renderActivityLegend(
  labels,
  values,
  colors
) {

  const legend =
    $("activityLegend");

  if (
    !legend
  ) {
    return;
  }

  const total =
    values.reduce(
      (sum, value) => sum + Number(value || 0),
      0
    );

  const pct = value =>
    total > 0
      ? ((Number(value || 0) / total) * 100).toFixed(1) + "%"
      : "0.0%";

  legend.innerHTML =
    labels.map(
      (
        label,
        index
      ) => `

        <button
          type="button"
          class="activity-legend-item"
          data-index="${index}"
        >

          <span
            class="activity-legend-dot"
            style="background:${colors[index]}"
          ></span>

          <span class="activity-legend-name">
            ${esc(label)}
          </span>

          <strong class="activity-legend-value">
            ${Number(values[index] || 0).toLocaleString()} (${pct(values[index])})
          </strong>

        </button>
      `
    )
      .join(
        ""
      );

  legend
    .querySelectorAll(
      ".activity-legend-item"
    )
    .forEach(
      item => {

        item.addEventListener(
          "click",
          () => {

            const index =
              Number(
                item.dataset
                  .index
              );

            if (
              labels[index]
            ) {

              openPopup(
                "type",
                labels[index]
              );
            }
          }
        );
      }
    );
}


/* =========================================================
   MONTHLY TREND
   ========================================================= */

function monthKey(value) {

  const text =
    clean(value);

  if (
    !text
  ) {
    return "";
  }

  if (
    /^\d{4}-\d{2}-\d{2}$/
      .test(
        text
      )
  ) {

    return text.slice(
      0,
      7
    );
  }

  if (
    /^\d{4}-\d{2}$/
      .test(
        text
      )
  ) {

    return text;
  }

  const date =
    new Date(
      text
    );

  if (
    !isNaN(
      date
    )
  ) {

    return date
      .toISOString()
      .slice(
        0,
        7
      );
  }

  return text;
}


function monthlyChart() {

  const canvas = $("monthChart");

  if (!canvas) {
    return;
  }

  destroyChart("monthChart");

  const result = {};

  filteredRows.forEach(row => {
    const semester = semesterFromMonth(row.month);

    if (!semester || isInvalidLabel(semester)) {
      return;
    }

    result[semester] =
      (result[semester] || 0) +
      rowAmount(row);
  });

  const semesterSortKey = label => {
    const match = String(label).match(/\b(20\d{2})\b/);
    const year = match ? Number(match[1]) : 0;
    const half = String(label).toLowerCase().includes("spring") ? 1 : 2;
    return year * 10 + half;
  };

  const labels = Object.keys(result).sort((a, b) => semesterSortKey(a) - semesterSortKey(b));

  const semesterTotal = Object.values(result).reduce((sum, value) => sum + Number(value || 0), 0);
  const semesterPercent = value =>
    semesterTotal > 0
      ? ((Number(value || 0) / semesterTotal) * 100).toFixed(1) + "%"
      : "0.0%";

  charts.monthChart =
    new Chart(
      canvas,
      {
        type: "line",
        data: {
          labels,
          datasets: [
            {
              data: labels.map(label => result[label]),
              borderColor: "#0b6b3a",
              backgroundColor: "#0b6b3a",
              borderWidth: 4,
              fill: false,
              tension: 0,
              spanGaps: true,
              pointRadius: 7,
              pointHoverRadius: 8,
              pointBackgroundColor: "#ffffff",
              pointBorderColor: "#0b6b3a",
              pointBorderWidth: 3
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: context => {
                  const value = Number(context.raw || 0);
                  return `${value.toLocaleString()} (${semesterPercent(value)})`;
                }
              }
            }
          },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { precision: 0 }
            },
            x: {
              offset: true,
              grid: { display: false },
              ticks: { autoSkip: false, font: { size: 11, weight: "bold" } }
            }
          }
        }
      }
    );
}


/* =========================================================
   RENDER
   ========================================================= */

function render() {

  const totalActivities =
    filteredRows.reduce(
      (
        sum,
        row
      ) =>
        sum +
        rowAmount(
          row
        ),
      0
    );

  const scheduled =
    filteredRows
      .filter(
        row =>
          yes(
            row.scheduled
          )
      )
      .reduce(
        (
          sum,
          row
        ) =>
          sum +
          rowAmount(
            row
          ),
        0
      );

  const planned =
    filteredRows
      .filter(
        row =>
          yes(
            row.planned
          )
      )
      .reduce(
        (
          sum,
          row
        ) =>
          sum +
          rowAmount(
            row
          ),
        0
      );

  const faculties =
    new Set(
      filteredRows
        .map(
          row =>
            row.faculty
        )
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    );

  const departments =
    new Set(
      filteredRows
        .map(
          row =>
            row.department
        )
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    );

  const types =
    new Set(
      filteredRows
        .map(
          row =>
            row.type
        )
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    );

  const setText =
    (
      id,
      value
    ) => {

      if (
        $(id)
      ) {

        $(id)
          .textContent =
          value;
      }
    };

  const selectedActivityType = $("activityFilter")?.value || "All";

  const totalActivitiesKpiCard = $("totalActivitiesKpiCard");
  if (totalActivitiesKpiCard) {
    // Show the overall total only when viewing all activity types.
    // Once a specific activity type is selected, this KPI is hidden.
    totalActivitiesKpiCard.style.display =
      selectedActivityType === "All" ? "" : "none";
  }

  const internshipKpiCard = $("internshipKpiCard");
  if (internshipKpiCard) {
    internshipKpiCard.style.display =
      selectedActivityType === "All" ||
      selectedActivityType === "Internship / Placement"
        ? ""
        : "none";
  }

  setText(
    "kActivities",
    totalActivities
      .toLocaleString()
  );
setText(
    "kFaculties",
    faculties
      .size
      .toLocaleString()
  );

  setText(
    "kDepartments",
    departments
      .size
      .toLocaleString()
  );

  setText(
    "kTypes",
    types
      .size
      .toLocaleString()
  );

  setText(
    "flowReported",
    totalActivities
      .toLocaleString()
  );

  setText(
    "flowPlanned",
    planned
      .toLocaleString()
  );
renderStudentSummary();

  barChart(
    "facultyChart",
    group(
      "faculty"
    ),
    "faculty",
    "facultyWrap"
  );

  barChart(
    "departmentChart",
    group(
      "department"
    ),
    "department",
    "departmentWrap"
  );

  activityChart();

  monthlyChart();
}


/* =========================================================
   WORKBOOK LOAD
   ========================================================= */

function findActivitySheet(
  workbook
) {

  const preferredNames =
    [
      "All Separated Activities",
      "Separated Activities",
      "Activities",
      "All Activities"
    ];

  for (
    const preferredName
    of preferredNames
  ) {

    const found =
      workbook
        .SheetNames
        .find(
          name =>
            norm(name) ===
            norm(
              preferredName
            )
        );

    if (
      found
    ) {
      return found;
    }
  }

  for (
    const sheetName
    of workbook.SheetNames
  ) {

    const rows =
      XLSX.utils
        .sheet_to_json(
          workbook
            .Sheets[
              sheetName
            ],
          {
            defval:
              "",
            raw:
              false
          }
        );

    if (
      !rows.length
    ) {
      continue;
    }

    const keys =
      Object.keys(
        rows[0]
      )
        .map(
          norm
        );

    if (
      keys.includes(
        "faculty"
      ) &&
      keys.includes(
        "department"
      )
    ) {

      return sheetName;
    }
  }

  return null;
}


async function fetchWorkbook(
  paths
) {

  let lastError =
    null;

  for (
    const path
    of paths
  ) {

    try {

      const response =
        await fetch(
          path,
          {
            cache:
              "no-store"
          }
        );

      if (
        !response.ok
      ) {

        throw new Error(
          `HTTP ${response.status}`
        );
      }

      const buffer =
        await response
          .arrayBuffer();

      return {

        path:
          path,

        workbook:
          XLSX.read(
            new Uint8Array(
              buffer
            ),
            {
              type:
                "array",
              cellDates:
                false
            }
          )
      };

    } catch (
      error
    ) {

      lastError =
        error;

      console.warn(
        "Workbook failed:",
        path,
        error
      );
    }
  }

  throw (
    lastError ||
    new Error(
      "Workbook not found"
    )
  );
}


async function loadActivityWorkbook() {

  const paths =
    [
      "data/UOL_Separated_All_Activity_Data_Cleaned.xlsx",
      "./data/UOL_Separated_All_Activity_Data_Cleaned.xlsx",
      "UOL_Separated_All_Activity_Data_Cleaned.xlsx",
      "./UOL_Separated_All_Activity_Data_Cleaned.xlsx"
    ];

  const result =
    await fetchWorkbook(
      paths
    );

  const workbook =
    result.workbook;

  const activitySheet =
    findActivitySheet(
      workbook
    );

  if (
    !activitySheet
  ) {

    throw new Error(
      "All Separated Activities sheet not found"
    );
  }

  const activityRows =
    XLSX.utils
      .sheet_to_json(
        workbook
          .Sheets[
            activitySheet
          ],
        {
          defval:
            "",
          raw:
            false
        }
      );

  allRows =
    parseActivities(
      activityRows
    );

  // Keep Remote activities as a separate dashboard category.
  // A record is treated as Remote when its source remarks explicitly say Mode: Remote.
  allRows = allRows.map(row => {
    const isRemote = /\bmode\s*:\s*remote\b/i.test(clean(row.remarks));
    return isRemote ? { ...row, type: "Remote" } : row;
  });

  // Planned is stored separately in the 56-submission workbook.
  // Add it to the dashboard as its own activity category instead of
  // mixing it with Internship / Placement.
  const plannedSheetName = workbook.SheetNames.find(name =>
    ["Planned Activities", "Planned - Section 4"].includes(name)
  );

  if (plannedSheetName) {
    const plannedRowsRaw = XLSX.utils.sheet_to_json(
      workbook.Sheets[plannedSheetName],
      { defval: "", raw: false }
    );

    const plannedRows = parseActivities(plannedRowsRaw)
      .map(row => {
        const isRemote = /\bmode\s*:\s*remote\b/i.test(clean(row.remarks));
        return { ...row, type: isRemote ? "Remote" : "Planned" };
      });

    const seen = new Set(allRows.map(row => [
      row.ref, row.department, row.originalType, row.events, row.organization, row.month
    ].join("||").toLowerCase()));

    plannedRows.forEach(row => {
      const key = [
        row.ref, row.department, row.originalType, row.events, row.organization, row.month
      ].join("||").toLowerCase();

      if (!seen.has(key)) {
        seen.add(key);
        allRows.push(row);
      }
    });
  }

  filteredRows =
    [...allRows];

  buildIpBreakdownFromActivities();

  console.log(
    "Workbook:",
    result.path
  );

  console.log(
    "Sheet:",
    activitySheet
  );

  console.log(
    "Records:",
    allRows.length
  );
}


async function loadExcel() {

  if (
    $("connection")
  ) {

    $("connection")
      .innerHTML =
      `<span style="color:#f59e0b">●</span> Loading workbook…`;
  }

  try {

    await loadActivityWorkbook();

    populateFilters();

    const requestedActivity = new URLSearchParams(window.location.search).get("activity");
    if (requestedActivity && $("activityFilter")) {
      const wanted = Array.from($("activityFilter").options)
        .find(option => norm(option.value) === norm(requestedActivity));
      if (wanted) {
        $("activityFilter").value = wanted.value;
        updateOrganizationFilter();
        applyFilters();
      }
    }

    render();

    const totalActivities =
      allRows.reduce(
        (
          sum,
          row
        ) =>
          sum +
          rowAmount(
            row
          ),
        0
      );

    const faculties =
      new Set(
        allRows.map(
          row =>
            row.faculty
        )
      );

    const departments =
      new Set(
        allRows.map(
          row =>
            row.department
        )
      );

    const ipSummary =
      summarizeIpRows(
        allIpRows
      );

    if (
      $("connection")
    ) {

      $("connection")
        .innerHTML = `
          <span style="color:#16a34a">●</span>
          Connected
          • ${allRows.length.toLocaleString()} records
          • ${totalActivities.toLocaleString()} activities
          <br>
          ${faculties.size.toLocaleString()} faculties
          • ${departments.size.toLocaleString()} departments
          • <strong>${ipSummary.total.toLocaleString()} Internship / Placement students</strong>
        `;
    }

  } catch (
    error
  ) {

    console.error(
      "Dashboard load error:",
      error
    );

    if (
      $("connection")
    ) {

      $("connection")
        .innerHTML =
        `<span style="color:#dc2626">●</span> Excel workbook failed`;
    }
  }
}


/* =========================================================
   POPUP
   ========================================================= */

function displayValue(value) {

  return isInvalidLabel(
    value
  )
    ? "—"
    : esc(
        value
      );
}


function openPopup(
  key,
  label
) {

  const modal =
    $("detailModal");

  if (
    !modal
  ) {
    return;
  }

  const rows =
    filteredRows.filter(
      row =>
        row[key] ===
        label
    );

  if (
    !rows.length
  ) {
    return;
  }

  if (
    $("detailModalTitle")
  ) {

    $("detailModalTitle")
      .textContent =
      label;
  }

  if (
    $("detailModalBody")
  ) {

    $("detailModalBody")
      .innerHTML =
      rows.map(
        (
          row,
          index
        ) => `

          <div class="detail-record">

            <strong>
              Record ${index + 1}
            </strong>

            <div class="detail-grid">

              <div class="detail-field">
                <small>Faculty</small>
                <strong>${displayValue(row.faculty)}</strong>
              </div>

              <div class="detail-field">
                <small>Department</small>
                <strong>${displayValue(row.department)}</strong>
              </div>

              <div class="detail-field">
                <small>Activity Type</small>
                <strong>${displayValue(row.type)}</strong>
              </div>

              <div class="detail-field">
                <small>Original Activity Type</small>
                <strong>${displayValue(row.originalType)}</strong>
              </div>

              <div class="detail-field">
                <small>How Many</small>
                <strong>${Number(row.events || 0).toLocaleString()}</strong>
              </div>

              <div class="detail-field">
                <small>Reporting Month</small>
                <strong>${displayValue(row.month)}</strong>
              </div>

              <div class="detail-field">
                <small>Person</small>
                <strong>${displayValue(row.person)}</strong>
              </div>

              <div class="detail-field">
                <small>Designation</small>
                <strong>${displayValue(row.designation)}</strong>
              </div>

              <div class="detail-field">
                <small>Organization</small>
                <strong>${displayValue(row.organization)}</strong>
              </div>

              <div class="detail-field">
                <small>Planned</small>
                <strong>${displayValue(row.planned)}</strong>
              </div>

              <div class="detail-field">
                <small>Scheduled</small>
                <strong>${displayValue(row.scheduled)}</strong>
              </div>

              <div class="detail-field">
                <small>Remarks / Result</small>
                <strong>${displayValue(row.remarks)}</strong>
              </div>

            </div>

          </div>
        `
      )
        .join(
          ""
        );
  }

  modal
    .classList
    .add(
      "open"
    );
}


function closePopup() {

  $("detailModal")
    ?.classList
    .remove(
      "open"
    );
}


/* =========================================================
   EXPORT
   ========================================================= */

function exportReport() {

  const workbook =
    XLSX.utils
      .book_new();

  const activityExport =
    filteredRows.map(
      row => ({

        "Ref #":
          row.ref,

        "Reporting Month":
          row.month,

        "Faculty":
          row.faculty,

        "Department":
          row.department,

        "Activity Type":
          row.type,

        "Original Activity Type":
          row.originalType,

        "How Many":
          row.events,

        "Focal Person":
          row.person,

        "Designation":
          row.designation,

        "Organization":
          row.organization,

        "Planned":
          row.planned,

        "Scheduled":
          row.scheduled,

        "Remarks / Result":
          row.remarks
      })
    );

  XLSX.utils
    .book_append_sheet(
      workbook,
      XLSX.utils
        .json_to_sheet(
          activityExport
        ),
      "Filtered Activities"
    );

  const ipRows =
    getFilteredIpRows();

  const ipExport =
    ipRows.map(
      row => ({

        "Faculty":
          row.faculty,

        "Department":
          row.department,

        "Type":
          row.sourceType,

        "Students":
          row.students
      })
    );

  XLSX.utils
    .book_append_sheet(
      workbook,
      XLSX.utils
        .json_to_sheet(
          ipExport
        ),
      "Internship Placement"
    );

  XLSX.writeFile(
    workbook,
    "UOL_Filtered_Activity_Report.xlsx"
  );
}


/* =========================================================
   EVENTS
   ========================================================= */

function bindEvents() {

  $("facultyFilter")
    ?.addEventListener(
      "change",
      handleFacultyChange
    );

  $("departmentFilter")
    ?.addEventListener(
      "change",
      handleDepartmentChange
    );

  $("activityFilter")
    ?.addEventListener(
      "change",
      handleActivityChange
    );

  $("monthFilter")
    ?.addEventListener(
      "change",
      handleMonthChange
    );

  $("organizationFilter")
    ?.addEventListener(
      "change",
      handleOrganizationChange
    );

  $("searchFilter")
    ?.addEventListener(
      "input",
      applyFilters
    );

  $("resetBtn")
    ?.addEventListener(
      "click",
      resetFilters
    );

  $("exportBtn")
    ?.addEventListener(
      "click",
      exportReport
    );

  if (
    $("mobileNavToggle") &&
    $("sidebar")
  ) {

    $("mobileNavToggle")
      .addEventListener(
        "click",
        () => {

          $("sidebar")
            .classList
            .toggle(
              "mobile-open"
            );
        }
      );
  }

  $("detailClose")
    ?.addEventListener(
      "click",
      closePopup
    );

  $("detailModal")
    ?.addEventListener(
      "click",
      event => {

        if (
          event.target ===
          $("detailModal")
        ) {

          closePopup();
        }
      }
    );

  document
    .addEventListener(
      "keydown",
      event => {

        if (
          event.key ===
          "Escape"
        ) {

          closePopup();
        }
      }
    );
}


/* =========================================================
   START
   ========================================================= */

document.addEventListener(
  "DOMContentLoaded",
  () => {

    bindEvents();

    loadExcel();
  }
);

/* =========================================================
   PDF / PRINT REPORTING
   Uses the same parsed/cleaned rows and dashboard filters.
   Report scopes:
   - current: current filtered dashboard view
   - faculties: one section for each faculty
   - departments: one section for each department
   - complete: faculty summary + each department under faculty
   Also includes graphs in PDF and print reports.
   ========================================================= */

function reportAmount(rows) {
  return rows.reduce((sum, row) => sum + rowAmount(row), 0);
}

function reportPercentage(value, total) {
  const n = Number(value) || 0;
  const t = Number(total) || 0;
  if (t <= 0) return "0.0%";
  return `${((n / t) * 100).toFixed(1)}%`;
}

function reportScheduled(rows) {
  return rows
    .filter(row => yes(row.scheduled))
    .reduce((sum, row) => sum + rowAmount(row), 0);
}

function reportPlanned(rows) {
  return rows
    .filter(row => yes(row.planned))
    .reduce((sum, row) => sum + rowAmount(row), 0);
}

function reportStudentTotal(rows) {
  return rows
    .filter(row => row.type === "Internship / Placement")
    .reduce((sum, row) => sum + rowAmount(row), 0);
}

function reportUnique(rows, key) {
  return new Set(
    rows.map(row => row[key]).filter(value => !isInvalidLabel(value))
  ).size;
}

function getCurrentFilterSummary() {
  const get = id => $(id)?.value || "All";
  return {
    faculty: get("facultyFilter"),
    department: get("departmentFilter"),
    activity: get("activityFilter"),
    month: get("monthFilter"),
    organization: get("organizationFilter"),
    search: $("searchFilter")?.value?.trim() || ""
  };
}


function reportFilterDetails(section = null) {
  const current = getCurrentFilterSummary();

  return [
    ["Semester", current.month === "All" ? "All Semesters" : current.month],
    ["Faculty", current.faculty === "All" ? "All Faculties" : current.faculty],
    ["Department", current.department === "All" ? "All Departments" : current.department],
    ["Activity Type", current.activity === "All" ? "All Activity Types" : current.activity],
    ["Organization", current.organization === "All" ? "All Organizations" : current.organization]
  ];
}

function addPdfFilterBanner(doc, section, startY = 57) {
  const width = doc.internal.pageSize.getWidth();
  const items = Object.fromEntries(reportFilterDetails(section));
  const left = 14;
  const right = 14;
  const gap = 3;
  const available = width - left - right;
  const boxW = (available - gap * 4) / 5;
  const boxH = 30;

  const boxes = [
    ["SEMESTER", items["Semester"], left, startY, boxW],
    ["FACULTY", items["Faculty"], left + (boxW + gap), startY, boxW],
    ["DEPARTMENT", items["Department"], left + (boxW + gap) * 2, startY, boxW],
    ["ACTIVITY TYPE", items["Activity Type"], left + (boxW + gap) * 3, startY, boxW],
    ["ORGANIZATION", items["Organization"], left + (boxW + gap) * 4, startY, boxW]
  ];

  boxes.forEach(([label, value, x, y, w]) => {
    doc.setFillColor(248, 250, 252);
    doc.setDrawColor(193, 211, 203);
    doc.setLineWidth(0.4);
    doc.roundedRect(x, y, w, boxH, 2.5, 2.5, "FD");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.2);
    doc.setTextColor(11, 107, 58);
    doc.text(label, x + w / 2, y + 5.2, { align: "center" });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.0);
    doc.setTextColor(0, 0, 0);
    const valueLines = doc.splitTextToSize(String(value || "-"), w - 6);
    const lines = valueLines.slice(0, 5);
    const lineHeight = 3.8;
    const totalTextHeight = Math.max(lineHeight, lines.length * lineHeight);
    const startTextY = y + 10 + Math.max(0, (boxH - 12 - totalTextHeight) / 2);

    doc.text(lines, x + w / 2, startTextY, {
      align: "center",
      lineHeightFactor: 1.25
    });
  });

  return startY + boxH + 7;
}

function reportSections() {
  const f = getCurrentFilterSummary();
  const parts = [];

  if (f.month !== "All") parts.push(f.month);
  if (f.faculty !== "All") parts.push(f.faculty);
  if (f.department !== "All") parts.push(f.department);
  if (f.activity !== "All") parts.push(f.activity);
  if (f.organization !== "All") parts.push(f.organization);
  if (f.search) parts.push(`Search: ${f.search}`);

  return [{
    title: parts.length ? parts.join(" - ") : "University-wide Activity Report",
    subtitle: "",
    rows: [...filteredRows],
    kind: "current"
  }];
}

function reportFileName(ext) {
  const faculty = $("facultyFilter")?.value || "All";
  const department = $("departmentFilter")?.value || "All";
  let name = "UOL_OEL_Report";

  if (faculty !== "All") name += `_${faculty}`;
  if (department !== "All") name += `_${department}`;

  return `${name.replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "")}.${ext}`;
}

function reportShouldShowInternship(rows) {
  const selectedActivity = $("activityFilter")?.value || "All";
  const types = unique(rows.map(row => row.type));

  if (selectedActivity === "Internship / Placement") {
    return true;
  }

  return types.length === 1 && types[0] === "Internship / Placement";
}

function reportMeta(rows) {
  const selectedActivity = $("activityFilter")?.value || "All";

  const meta = [
    ["Faculties", reportUnique(rows, "faculty").toLocaleString()],
    ["Departments", reportUnique(rows, "department").toLocaleString()],
    ["Activity Types", reportUnique(rows, "type").toLocaleString()],
    ["Organizations", reportUnique(rows, "organization").toLocaleString()]
  ];

  if (selectedActivity === "All") {
    meta.unshift(["Total Activities", reportAmount(rows).toLocaleString()]);
  }

  if (reportShouldShowInternship(rows)) {
    meta.splice(1, 0, [
      "Internship / Placement Students",
      reportStudentTotal(rows).toLocaleString()
    ]);
  }

  return meta;
}

const REPORT_LOGO_PATH = "assets/uol-wordmark.png";
let reportLogoCache = "";

async function getReportLogoDataUrl() {
  if (reportLogoCache) return reportLogoCache;

  try {
    const response = await fetch(REPORT_LOGO_PATH);
    const blob = await response.blob();
    reportLogoCache = await new Promise(resolve => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    reportLogoCache = "";
  }

  return reportLogoCache;
}

function reportGeneratedOn() {
  return new Date().toLocaleString();
}

function focalPersonBreakdown(rows) {
  const grouped = {};

  rows.forEach(row => {
    const name = clean(row.person);
    if (isInvalidLabel(name)) return;

    const designation = clean(row.designation) || "-";
    const department = clean(row.department) || "-";
    const key = `${name}||${designation}||${department}`;

    if (!grouped[key]) {
      grouped[key] = {
        name,
        designation,
        department,
        total: 0
      };
    }

    grouped[key].total += rowAmount(row);
  });

  return Object.values(grouped)
    .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
}

function reportMetaTableRows(meta) {
  const rows = [];
  for (let i = 0; i < meta.length; i += 2) {
    const left = meta[i] || ["", ""];
    const right = meta[i + 1] || ["", ""];
    rows.push([left[0], left[1], right[0], right[1]]);
  }
  return rows;
}


function reportMetaRows(meta) {
  const rows = [];
  for (let i = 0; i < meta.length; i += 2) {
    const a = meta[i] || ["", ""];
    const b = meta[i + 1] || ["", ""];
    rows.push([a[0], a[1], b[0], b[1]]);
  }
  return rows;
}

function activityBreakdown(rows) {
  const grouped = {};
  rows.forEach(row => {
    const key = row.type || "Other";
    grouped[key] = (grouped[key] || 0) + rowAmount(row);
  });
  return Object.entries(grouped).sort((a, b) => b[1] - a[1]);
}

function reportDisplayValue(value) {
  const cleaned = clean(value);
  return !cleaned || isInvalidLabel(cleaned) ? "-" : cleaned;
}

function detailRows(rows) {
  return rows.map(row => [
    semesterFromMonth(row.month) || "-",
    reportDisplayValue(row.faculty),
    reportDisplayValue(row.department),
    reportDisplayValue(row.type),
    rowAmount(row),
    reportDisplayValue(row.organization),
    reportDisplayValue(row.person),
    reportDisplayValue(row.designation),
    reportDisplayValue(row.scheduled)
  ]);
}

function expectedOutcomeBullets(rows) {
  const present = new Set(rows.map(row => clean(row.type)).filter(Boolean));
  const rules = [
    ["Internship / Placement", "Internship / Placement: Improve student employability and workplace readiness by expanding structured opportunities for practical training, professional exposure and recruitment. Expected results include stronger employer networks, improved job-readiness, better understanding of workplace standards and a clearer pathway from academic learning to employment."],
    ["Planned", "Planned Activities: Establish a forward-looking pipeline of external-linkage opportunities with clear partner organizations, student participation targets, timelines, delivery modes and responsible departments. Expected results include improved semester planning, earlier coordination with external partners, better conversion of planned opportunities into executed internships, placements, visits or collaborations, and stronger evidence-based monitoring of progress."],
    ["Remote", "Remote Activities: Expand access to external-linkage opportunities through virtual or remote participation where physical attendance is not required. Expected results include broader student participation, flexible engagement with external organizations, improved accessibility, documented remote delivery, and measurable academic or professional outcomes from online placements, collaborations or planned activities."],
    ["Alumni Talk / Mentoring", "Alumni Talk / Mentoring: Strengthen alumni participation in student development through career guidance, mentoring, networking and sharing of professional experience. Expected results include improved career awareness, stronger alumni-student connections and access to sector-specific advice and opportunities."],
    ["Industrial Visit / IV", "Industrial Visit / IV: Provide direct exposure to operational environments, technologies, professional practices and organizational systems. Expected results include improved understanding of industry processes, stronger application of classroom learning and increased awareness of workplace expectations."],
    ["Seminar", "Seminar: Enhance knowledge exchange by connecting students and faculty with external experts, practitioners and current developments in relevant disciplines. Expected results include improved professional awareness, updated subject knowledge and opportunities for academic and industry networking."],
    ["Guest Lecture / GLIT", "Guest Lecture / GLIT: Integrate industry and external-expert perspectives into teaching and learning. Expected results include stronger understanding of real-world practices, emerging trends, career pathways and professional competencies required by employers."],
    ["Research Collaboration", "Research Collaboration: Expand joint research, interdisciplinary projects, access to external expertise and opportunities for collaborative publications or funded proposals. Expected results include stronger research networks, improved research impact and greater institutional visibility."],
    ["MoU / MoU Signing", "MoU / MoU Signing: Formalize sustainable relationships with academic institutions, industry, government, NGOs and international partners. Expected results include a structured framework for internships, research, faculty/student exchange, training, joint events and future collaborative initiatives."],
    ["Curriculum Feedback", "Curriculum Feedback: Incorporate employer, professional and external stakeholder feedback into academic planning. Expected results include improved curriculum relevance, stronger alignment with labour-market needs and identification of skill gaps that can be addressed through programme improvement."],
    ["Community Engagement", "Community Engagement: Increase the university's contribution to society through outreach, awareness, service and collaborative community initiatives. Expected results include stronger social impact, improved stakeholder relationships and increased opportunities for experiential learning."],
    ["Workshop", "Workshop: Build practical and professional capacity through hands-on training, demonstrations and applied learning. Expected results include development of technical and soft skills, improved confidence in applying knowledge and exposure to current tools and professional practices."],
    ["IAB / Industry Consultation", "IAB / Industry Consultation: Strengthen systematic industry participation in academic and strategic planning. Expected results include actionable employer input, programme improvement, stronger university-industry coordination and better alignment of graduate competencies with sector requirements."],
    ["Conference", "Conference: Promote scholarly exchange, dissemination of research, professional networking and institutional visibility. Expected results include exposure to current research and practice, development of academic networks and opportunities for future collaborative research and professional engagement."],
    ["Others", "Other External-Linkage Activities: Support additional initiatives that contribute to institutional engagement, external visibility, partnership development or student and faculty development. Expected outcomes should be assessed against the specific purpose and measurable results of each activity."]
  ];

  const bullets = rules.filter(([type]) => present.has(type)).map(([, outcome]) => outcome);
  if (!bullets.length) {
    bullets.push("Overall Expected Outcome: Strengthen external engagement and convert planned activities into measurable academic, professional and institutional outcomes through sustained partnerships, documented implementation and evidence of impact.");
  }
  return bullets;
}

function reportManagementRemarks(rows) {
  const total = rows.reduce((sum, row) => sum + rowAmount(row), 0);
  const activity = activityBreakdown(rows);
  const departments = new Set(rows.map(row => clean(row.department)).filter(Boolean));
  const faculties = new Set(rows.map(row => clean(row.faculty)).filter(Boolean));
  const remarks = [];

  if (total > 0) {
    remarks.push(`Overall review: The selected report contains ${total.toLocaleString()} reported activity outcomes across ${departments.size} department${departments.size === 1 ? "" : "s"} and ${faculties.size} facult${faculties.size === 1 ? "y" : "ies"}. The figures should be reviewed together with supporting evidence and completion status to distinguish planned activity from executed and verified outcomes.`);
  }

  if (activity.length) {
    const [topType, topValue] = activity[0];
    const pct = total ? ((topValue / total) * 100).toFixed(1) : "0.0";
    remarks.push(`Activity concentration: ${topType} is the largest activity category with ${topValue.toLocaleString()} reported outcomes (${pct}% of the selected total). OEL may review whether the portfolio is sufficiently balanced across employability, partnerships, research collaboration, industry consultation, curriculum engagement and other strategic linkage areas.`);
  }

  const mou = activity.find(([type]) => type === "MoU / MoU Signing");
  if (mou) {
    remarks.push(`Partnership follow-through: ${mou[1].toLocaleString()} MoU / MoU Signing outcome${mou[1] === 1 ? " is" : "s are"} reported. Departments should link each formal partnership with subsequent activities, responsible persons, timelines, evidence and measurable benefits so that signed agreements translate into active collaboration.`);
  }

  const internship = activity.find(([type]) => type === "Internship / Placement");
  if (internship) {
    remarks.push(`Employability follow-through: ${internship[1].toLocaleString()} Internship / Placement outcome${internship[1] === 1 ? " is" : "s are"} reported. Departments should maintain organization-wise evidence, student participation records, completion status and, where possible, conversion to employment or other measurable career outcomes.`);
  }

  const research = activity.find(([type]) => type === "Research Collaboration");
  if (research) {
    remarks.push(`Research follow-through: ${research[1].toLocaleString()} Research Collaboration outcome${research[1] === 1 ? " is" : "s are"} reported. Future monitoring should capture resulting proposals, publications, grants, shared facilities, researcher exchanges or other tangible research outputs.`);
  }

  remarks.push("Monitoring recommendation: For the next review cycle, departments should report progress against semester plans using clear status categories (planned, scheduled, executed and completed), attach supporting evidence and state the actual outcome achieved for students, faculty, partners or the institution.");

  return remarks;
}

function reportRemarksBullets(rows) {
  const seen = new Set();
  const remarks = [];

  rows.forEach(row => {
    const value = normalizeDisplayText(row.remarks);
    if (!value || isInvalidLabel(value)) return;
    const normalized = value.toLowerCase().replace(/\s+/g, " ").trim();
    if (!normalized || seen.has(normalized)) return;
    seen.add(normalized);
    remarks.push(value);
  });

  return remarks.slice(0, 18);
}

function drawPdfBulletSection(doc, title, bullets, startY, options = {}) {
  const left = options.left || 18;
  const right = options.right || 18;
  const width = doc.internal.pageSize.getWidth();
  const maxWidth = width - left - right - 8;
  let y = startY;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(12.5);
  doc.setTextColor(9, 31, 84);
  doc.text(title, left, y);
  y += 7;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.1);
  doc.setTextColor(15, 23, 42);

  bullets.forEach(text => {
    const lines = doc.splitTextToSize(String(text), maxWidth);
    doc.setFont("helvetica", "bold");
    doc.text("•", left + 1, y);
    doc.setFont("helvetica", "normal");
    doc.text(lines, left + 6, y, { lineHeightFactor: 1.35 });
    y += Math.max(6, lines.length * 4.25 + 2.5);
  });

  return y;
}

function addPdfExpectedOutcomesPage(doc, section, sectionIndex, sectionCount, logoDataUrl = "") {
  const outcomes = expectedOutcomeBullets(section.rows);

  const pageSection = {
    ...section,
    subtitle: `${section.subtitle || ""} - Detailed Expected Outcomes`
  };

  const pageHeight = doc.internal.pageSize.getHeight();
  const bottomLimit = pageHeight - 28;

  const newContinuationPage = () => {
    doc.addPage();
    addPdfHeader(doc, pageSection, sectionIndex, sectionCount, "EXPECTED OUTCOMES", logoDataUrl);
    return 60;
  };

  const drawPagedBullets = (title, bullets, y) => {
    if (y > bottomLimit - 25) y = newContinuationPage();

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12.5);
    doc.setTextColor(9, 31, 84);
    doc.text(title, 18, y);
    y += 7;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.4);
    doc.setTextColor(0, 0, 0);
    const maxWidth = doc.internal.pageSize.getWidth() - 44;

    bullets.forEach(text => {
      const lines = doc.splitTextToSize(String(text), maxWidth);
      const needed = Math.max(7, lines.length * 4.25 + 3);
      if (y + needed > bottomLimit) {
        y = newContinuationPage();
      }
      doc.setFont("helvetica", "bold");
      doc.text("•", 19, y);
      doc.setFont("helvetica", "normal");
      doc.text(lines, 24, y, { lineHeightFactor: 1.35 });
      y += needed;
    });
    return y;
  };

  addPdfHeader(doc, pageSection, sectionIndex, sectionCount, "EXPECTED OUTCOMES", logoDataUrl);
  drawPagedBullets("Detailed Expected Outcomes", outcomes, 60);
}

function addPdfRemarksPage(doc, section, sectionIndex, sectionCount, logoDataUrl = "") {
  const managementRemarks = reportManagementRemarks(section.rows);
  const recordedRemarks = reportRemarksBullets(section.rows);

  const pageSection = {
    ...section,
    subtitle: `${section.subtitle || ""} - Detailed Remarks`
  };

  const pageHeight = doc.internal.pageSize.getHeight();
  const bottomLimit = pageHeight - 28;

  const newContinuationPage = () => {
    doc.addPage();
    addPdfHeader(doc, pageSection, sectionIndex, sectionCount, "REMARKS", logoDataUrl);
    return 60;
  };

  const drawPagedBullets = (title, bullets, y) => {
    if (y > bottomLimit - 25) y = newContinuationPage();

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12.5);
    doc.setTextColor(9, 31, 84);
    doc.text(title, 18, y);
    y += 7;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9.4);
    doc.setTextColor(0, 0, 0);
    const maxWidth = doc.internal.pageSize.getWidth() - 44;

    bullets.forEach(text => {
      const lines = doc.splitTextToSize(String(text), maxWidth);
      const needed = Math.max(7, lines.length * 4.25 + 3);
      if (y + needed > bottomLimit) {
        y = newContinuationPage();
      }
      doc.setFont("helvetica", "bold");
      doc.text("•", 19, y);
      doc.setFont("helvetica", "normal");
      doc.text(lines, 24, y, { lineHeightFactor: 1.35 });
      y += needed;
    });
    return y;
  };

  addPdfHeader(doc, pageSection, sectionIndex, sectionCount, "REMARKS", logoDataUrl);

  let y = drawPagedBullets("Management Review Remarks", managementRemarks, 60);
  y += 5;

  const actualRemarks = recordedRemarks.length
    ? recordedRemarks
    : ["No specific Remarks / Result text was recorded in the filtered source data. This is a source-data gap and should be completed by the responsible department or focal person where applicable."];
  drawPagedBullets("Recorded Remarks / Results", actualRemarks, y);
}


function addPdfLogo(doc, logoDataUrl, x, y, maxWidth, maxHeight) {
  if (!logoDataUrl) return;
  // Official UOL wordmark aspect ratio from the supplied logo asset.
  const ratio = 793 / 336;
  let width = maxWidth;
  let height = width / ratio;
  if (height > maxHeight) {
    height = maxHeight;
    width = height * ratio;
  }
  try {
    doc.addImage(logoDataUrl, "PNG", x, y, width, height, undefined, "FAST");
  } catch (error) {}
  return { width, height };
}

function addPdfCoverPage(doc, sections, logoDataUrl = "") {
  const width = doc.internal.pageSize.getWidth();
  const height = doc.internal.pageSize.getHeight();
doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, width, height, "F");
  doc.setFillColor(9, 31, 84);
  doc.rect(0, 0, width, 13, "F");
  doc.setFillColor(11, 107, 58);
  doc.rect(0, 13, width, 4, "F");

  // Keep the official wordmark undistorted and let it carry the university name.
  const logoW = 92;
  const logoH = logoW / (793 / 336);
  addPdfLogo(doc, logoDataUrl, (width - logoW) / 2, 27, logoW, logoH);

  doc.setTextColor(11, 107, 58);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text("OFFICE OF EXTERNAL LINKAGES", width / 2, 77, { align: "center" });

  doc.setDrawColor(11, 107, 58);
  doc.setLineWidth(0.8);
  doc.line(width / 2 - 62, 83, width / 2 + 62, 83);

  doc.setTextColor(9, 31, 84);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(22);
  doc.text("External Linkages Intelligence Management System", width / 2, 97, { align: "center" });
  const cardX = 40;
  const cardW = width - 80;

  // Report leadership / authority block
  const authorityY = 116;
  const authorityGap = 7;
  const authorityW = (cardW - authorityGap) / 2;

  doc.setFillColor(246, 251, 248);
  doc.setDrawColor(189, 220, 202);
  doc.roundedRect(cardX, authorityY, authorityW, 13, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.6);
  doc.setTextColor(11, 107, 58);
  doc.text("DEAN OSA", cardX + 4, authorityY + 5);
  doc.setFontSize(9.2);
  doc.setTextColor(15, 23, 42);
  doc.text("Ms. Ammara Awais Raoof", cardX + 4, authorityY + 10.2);

  const authorityX2 = cardX + authorityW + authorityGap;
  doc.setFillColor(246, 249, 253);
  doc.setDrawColor(193, 207, 226);
  doc.roundedRect(authorityX2, authorityY, authorityW, 13, 2.5, 2.5, "FD");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.6);
  doc.setTextColor(9, 31, 84);
  doc.text("DIRECTOR OEL", authorityX2 + 4, authorityY + 5);
  doc.setFontSize(9.2);
  doc.setTextColor(15, 23, 42);
  doc.text("Dr. Muhammad Shafique", authorityX2 + 4, authorityY + 10.2);

  // Show the selected report filters only on the first PDF page.
  if (sections && sections.length) {
    addPdfFilterBanner(doc, sections[0], 141);
  }

  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.5);
  doc.setTextColor(0, 0, 0);
  doc.text(`Generated on: ${reportGeneratedOn()}`, width / 2, 180, { align: "center" });

  doc.setFillColor(9, 31, 84);
  doc.rect(0, height - 18, width, 18, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.5);
  doc.text("CONNECT - COLLABORATE - CREATE IMPACT", width / 2, height - 7, { align: "center" });
}

function addFacultyDividerPage(doc, section, logoDataUrl = "") {
  const width = doc.internal.pageSize.getWidth();
  const height = doc.internal.pageSize.getHeight();

  doc.setFillColor(247, 250, 252);
  doc.rect(0, 0, width, height, "F");
  doc.setFillColor(9, 31, 84);
  doc.rect(0, 0, width, 33, "F");
  doc.setFillColor(11, 107, 58);
  doc.rect(0, 33, width, 4, "F");

  addPdfLogo(doc, logoDataUrl, 16, 5, 58, 24);
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("OFFICE OF EXTERNAL LINKAGES", 82, 16);
  doc.setFontSize(8.5);
  doc.text("FACULTY ACTIVITY REPORT", 82, 24);

  doc.setTextColor(11, 107, 58);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text("FACULTY REPORT", 20, 68);

  doc.setTextColor(9, 31, 84);
  doc.setFontSize(23);
  const titleLines = doc.splitTextToSize(section.title || "Faculty", width - 40);
  doc.text(titleLines, 20, 84);

  const titleExtra = Math.max(0, titleLines.length - 1) * 9;
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(11);
  doc.text("Faculty-wise activity summary and detailed report", 20, 106 + titleExtra);

  const y = 124 + titleExtra;
  const facultyRows = section.rows || [];
  const metrics = [
    ["Departments", reportUnique(facultyRows, "department")],
    ["Activity Types", reportUnique(facultyRows, "type")],
    ["Scheduled", reportScheduled(facultyRows)]
  ];
  const cardW = 76;
  const gap = 9;
  metrics.forEach((item, i) => {
    const x = 20 + i * (cardW + gap);
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(207, 224, 216);
    doc.roundedRect(x, y, cardW, 30, 3, 3, "FD");
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.text(item[0].toUpperCase(), x + 5, y + 9);
    doc.setTextColor(9, 31, 84);
    doc.setFontSize(16);
    doc.text(Number(item[1] || 0).toLocaleString(), x + 5, y + 22);
  });
}

function addPdfHeader(doc, section, sectionIndex, sectionCount, label = "ACTIVITY REPORT", logoDataUrl = "") {
  const pageWidth = doc.internal.pageSize.getWidth();

  // IMPORTANT: only paint the header band here. autoTable calls this function
  // from didDrawPage after table rows have been rendered; painting the whole
  // page here would erase the Faculty, Activity and Department tables.
  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, pageWidth, 35, "F");
  doc.setFillColor(11, 107, 58);
  doc.rect(0, 31.5, pageWidth, 3.5, "F");
  doc.setDrawColor(219, 228, 240);
  doc.line(0, 35, pageWidth, 35);

  addPdfLogo(doc, logoDataUrl, 14, 4.5, 62, 25);

  doc.setTextColor(9, 31, 84);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12.5);
  doc.text("OFFICE OF EXTERNAL LINKAGES", 84, 13);
  doc.setFontSize(8.5);
  doc.setTextColor(0, 0, 0);
  doc.text("External Linkages Intelligence Management System", 84, 20.5);
  doc.setFontSize(7.2);
  doc.setTextColor(0, 0, 0);
  doc.text("Dean OSA: Ms. Ammara Awais Raoof  |  Director OEL: Dr. Muhammad Shafique", 84, 26.5);

  doc.setFillColor(9, 31, 84);
  doc.roundedRect(pageWidth - 75, 7, 61, 15, 3, 3, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.8);
  doc.text(label, pageWidth - 44.5, 16.5, { align: "center" });

  doc.setTextColor(15, 23, 42);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  const title = doc.splitTextToSize(section.title || "Report", 190);
  doc.text(title.slice(0, 1), 14, 45);

  doc.setFontSize(8.7);
  doc.setTextColor(0, 0, 0);
  doc.text(section.subtitle || "", 14, 51);

  doc.setFontSize(8.2);
  doc.setTextColor(0, 0, 0);
  doc.text(`Section ${sectionIndex + 1} of ${sectionCount}`, pageWidth - 14, 44, { align: "right" });
  doc.text(`Generated: ${reportGeneratedOn()}`, pageWidth - 14, 50.5, { align: "right" });
}


function addPdfFooter(doc) {
  const pages = doc.internal.getNumberOfPages();
  for (let page = 2; page <= pages; page++) {
    doc.setPage(page);
    const width = doc.internal.pageSize.getWidth();
    const height = doc.internal.pageSize.getHeight();
    doc.setDrawColor(219, 228, 240);
    doc.line(14, height - 14, width - 14, height - 14);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.8);
    doc.setTextColor(0, 0, 0);
    doc.text("THE UNIVERSITY OF LAHORE - Office of External Linkages", 14, height - 9.5);
    doc.setFontSize(6.8);
    doc.setTextColor(0, 0, 0);
    doc.text("Dean OSA: Ms. Ammara Awais Raoof  |  Director OEL: Dr. Muhammad Shafique", 14, height - 5.5);
    doc.setFontSize(7.8);
    doc.setTextColor(0, 0, 0);
    doc.text(`Page ${page - 1} of ${pages - 1}`, width - 14, height - 8, { align: "right" });
  }
}

function reportGroupBy(rows, key) {
  const grouped = {};

  rows.forEach(row => {
    const label = clean(row[key]);
    if (isInvalidLabel(label)) return;
    grouped[label] = (grouped[label] || 0) + rowAmount(row);
  });

  return Object.entries(grouped).sort((a, b) => b[1] - a[1]);
}

function reportMonthBreakdown(rows) {
  const grouped = {};

  rows.forEach(row => {
    const label = semesterFromMonth(row.month);
    if (!label || isInvalidLabel(label)) return;
    grouped[label] = (grouped[label] || 0) + rowAmount(row);
  });

  const sortKey = label => {
    const match = String(label).match(/\b(20\d{2})\b/);
    const year = match ? Number(match[1]) : 0;
    const half = String(label).toLowerCase().includes("spring") ? 1 : 2;
    return year * 10 + half;
  };

  return Object.entries(grouped).sort((a, b) => sortKey(a[0]) - sortKey(b[0]));
}

function reportChartPalette() {
  return [
    "#0b6b3a",
    "#1556c0",
    "#7f2db4",
    "#ef7d17",
    "#0b7f90",
    "#14206d",
    "#e04b5a",
    "#d4a017",
    "#166534",
    "#2563eb",
    "#9333ea",
    "#ea580c"
  ];
}

function chartBackgroundPlugin() {
  return {
    id: "reportChartBackground",
    beforeDraw(chart) {
      const { ctx, width, height } = chart;
      ctx.save();
      ctx.globalCompositeOperation = "source-over";
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, width, height);
      ctx.restore();
    }
  };
}

function semesterValueLabelsPlugin() {
  return {
    id: "semesterValueLabels",
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const dataset = chart.data.datasets[0];
      const meta = chart.getDatasetMeta(0);

      ctx.save();
      ctx.fillStyle = "#0f172a";
      ctx.font = "bold 34px Arial";
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";

      meta.data.forEach((point, index) => {
        const value = dataset.data[index];
        ctx.fillText(String(value), point.x, point.y - 22);
      });
      ctx.restore();
    }
  };
}

function wrapChartLabel(text, maxChars = 22) {
  const value = String(text || "").trim();
  if (!value) return ["-"];
  const words = value.split(/\s+/);
  const lines = [];
  let current = "";

  words.forEach(word => {
    const next = current ? `${current} ${word}` : word;
    if (next.length <= maxChars) {
      current = next;
    } else {
      if (current) lines.push(current);
      if (word.length > maxChars) {
        const chunks = word.match(new RegExp(`.{1,${maxChars}}`, "g")) || [word];
        if (chunks.length > 1) {
          lines.push(...chunks.slice(0, -1));
          current = chunks[chunks.length - 1];
        } else {
          current = word;
        }
      } else {
        current = word;
      }
    }
  });

  if (current) lines.push(current);
  return lines.slice(0, 3);
}

async function renderReportChart(config, width = 3200, height = 1600) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext("2d", { alpha: false });
  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.restore();
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";

  const mergedConfig = {
    ...config,
    plugins: [chartBackgroundPlugin(), ...(config.plugins || [])],
    options: {
      responsive: false,
      maintainAspectRatio: false,
      animation: false,
      devicePixelRatio: 1,
      layout: {
        padding: { left: 40, right: 40, top: 56, bottom: 24 },
        ...(config.options?.layout || {})
      },
      ...config.options
    }
  };

  const chart = new Chart(ctx, mergedConfig);
  chart.resize(width, height);
  chart.update("none");
  await new Promise(resolve => setTimeout(resolve, 100));
  const image = canvas.toDataURL("image/png");
  chart.destroy();
  return image;
}

async function buildSectionCharts(section) {
  const rows = section.rows || [];
  if (!rows.length) return [];

  const palette = reportChartPalette();
  const charts = [];
  const facultyEntries = reportGroupBy(rows, "faculty").slice(0, 10);
  const departmentEntries = reportGroupBy(rows, "department");
  const typeEntries = activityBreakdown(rows);
  const monthEntries = reportMonthBreakdown(rows);

  if (facultyEntries.length > 1) {
    charts.push({
      title: "Faculty-wise Activities",
      image: await renderReportChart({
        type: "bar",
        data: {
          labels: facultyEntries.map(item => wrapChartLabel(item[0], 24)),
          datasets: [{
            label: "Activities",
            data: facultyEntries.map(item => item[1]),
            backgroundColor: facultyEntries.map((_, index) => palette[index % palette.length]),
            borderRadius: 5,
            maxBarThickness: 72
          }]
        },
        options: {
          indexAxis: "y",
          plugins: { legend: { display: false }, title: { display: false } },
          scales: {
            x: { beginAtZero: true, ticks: { precision: 0, autoSkip: false, font: { size: 28, weight: "bold" }, color: "#0f172a", padding: 10 }, grid: { color: "rgba(15,23,42,.16)", lineWidth: 2 }, border: { color: "#334155", width: 3 } },
            y: { ticks: { autoSkip: false, font: { size: 28, weight: "bold" }, color: "#0f172a", padding: 12 }, grid: { display: false }, border: { color: "#334155", width: 3 } }
          }
        }
      }, 3600, 1650)
    });
  }

  if (departmentEntries.length > 1) {
    charts.push({
      title: "Department-wise Activities",
      image: await renderReportChart({
        type: "bar",
        data: {
          labels: departmentEntries.map(item => wrapChartLabel(item[0], 24)),
          datasets: [{
            label: "Activities",
            data: departmentEntries.map(item => item[1]),
            backgroundColor: departmentEntries.map((_, index) => palette[index % palette.length]),
            borderRadius: 5,
            maxBarThickness: 68
          }]
        },
        options: {
          indexAxis: "y",
          plugins: { legend: { display: false }, title: { display: false } },
          scales: {
            x: { beginAtZero: true, ticks: { precision: 0, autoSkip: false, font: { size: 28, weight: "bold" }, color: "#0f172a", padding: 10 }, grid: { color: "rgba(15,23,42,.16)", lineWidth: 2 }, border: { color: "#334155", width: 3 } },
            y: { ticks: { autoSkip: false, font: { size: 20, weight: "bold" }, color: "#0f172a", padding: 12 }, grid: { display: false }, border: { color: "#334155", width: 3 } }
          }
        }
      }, 3600, 3000)
    });
  }

  if (typeEntries.length) {
    charts.push({
      title: "Activity Type Distribution",
      image: await renderReportChart({
        type: "doughnut",
        data: {
          labels: typeEntries.map(item => item[0]),
          datasets: [{
            data: typeEntries.map(item => item[1]),
            backgroundColor: typeEntries.map((_, index) => palette[index % palette.length]),
            borderColor: "#ffffff",
            borderWidth: 5,
            hoverOffset: 4
          }]
        },
        options: {
          cutout: "55%",
          plugins: {
            title: { display: false },
            legend: { position: "right", labels: { boxWidth: 26, padding: 16, color: "#0f172a", font: { size: 24, weight: "bold" } } }
          }
        }
      }, 2400, 2200)
    });
  }

  if (monthEntries.length) {
    charts.push({
      title: "Semester Trend",
      image: await renderReportChart({
        type: "line",
        data: {
          labels: monthEntries.map(item => wrapChartLabel(item[0], 18)),
          datasets: [{
            label: "Activities",
            data: monthEntries.map(item => item[1]),
            borderColor: "#0b6b3a",
            backgroundColor: "#0b6b3a",
            borderWidth: 12,
            fill: false,
            tension: 0,
            spanGaps: true,
            pointRadius: 18,
            pointHoverRadius: 18,
            pointBackgroundColor: "#ffffff",
            pointBorderColor: "#0b6b3a",
            pointBorderWidth: 8
          }]
        },
        plugins: [semesterValueLabelsPlugin()],
        options: {
          plugins: {
            title: { display: false },
            legend: { display: false }
          },
          layout: { padding: { left: 80, right: 80, top: 80, bottom: 55 } },
          scales: {
            y: { beginAtZero: true, suggestedMax: Math.max(...monthEntries.map(item => item[1])) * 1.18, ticks: { precision: 0, autoSkip: false, font: { size: 30, weight: "bold" }, color: "#0f172a", padding: 12 }, grid: { color: "rgba(15,23,42,.14)", lineWidth: 2 }, border: { color: "#334155", width: 3 } },
            x: { offset: true, grid: { display: false }, ticks: { autoSkip: false, maxRotation: 0, minRotation: 0, color: "#0f172a", padding: 18, font: { size: 30, weight: "bold" } }, border: { color: "#334155", width: 3 } }
          }
        }
      }, 3200, 1650)
    });
  }

  return charts.slice(0, 4);
}


function addPdfImageContain(doc, imageData, x, y, w, h) {
  try {
    const props = doc.getImageProperties(imageData);
    const ratio = props.width / props.height;
    let drawW = w;
    let drawH = drawW / ratio;

    if (drawH > h) {
      drawH = h;
      drawW = drawH * ratio;
    }

    const drawX = x + (w - drawW) / 2;
    const drawY = y + (h - drawH) / 2;
    doc.addImage(imageData, "PNG", drawX, drawY, drawW, drawH, undefined, "NONE");
  } catch (error) {
    doc.addImage(imageData, "PNG", x, y, w, h, undefined, "NONE");
  }
}

function drawPdfDashboardGraphPair(doc, leftItem, rightItem, section, startY = 84) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 12;
  const gap = 5;
  const footerTop = pageHeight - 18;
  const availableHeight = footerTop - startY;

  const leftWidth = (pageWidth - margin * 2 - gap) * 0.66;
  const rightWidth = (pageWidth - margin * 2 - gap) - leftWidth;

  function drawCard(item, x, y, w, h, subtitle) {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(210, 218, 229);
    doc.setLineWidth(0.35);
    doc.roundedRect(x, y, w, h, 2.5, 2.5, "FD");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10.5);
    doc.setTextColor(9, 31, 84);
    doc.text((item?.title || "Graph").toUpperCase(), x + 3, y + 6);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.3);
    doc.setTextColor(0, 0, 0);
    doc.text(subtitle, x + 3, y + 11);

    const imageTop = y + 13;
    const imageBottom = y + h - 3;
    addPdfImageContain(
      doc,
      item.image,
      x + 3,
      imageTop,
      w - 6,
      imageBottom - imageTop
    );
  }

  if (leftItem && rightItem) {
    drawCard(
      leftItem,
      margin,
      startY,
      leftWidth,
      availableHeight,
      leftItem.title === "Faculty-wise Activities"
        ? "Total activity volume by faculty"
        : "Total activity volume by department"
    );

    drawCard(
      rightItem,
      margin + leftWidth + gap,
      startY,
      rightWidth,
      availableHeight,
      rightItem.title === "Activity Type Distribution"
        ? "Combined activity categories"
        : "Activities by semester"
    );
  } else {
    const onlyItem = leftItem || rightItem;
    drawCard(
      onlyItem,
      margin,
      startY,
      pageWidth - margin * 2,
      availableHeight,
      onlyItem?.title === "Faculty-wise Activities"
        ? "Total activity volume by faculty"
        : onlyItem?.title === "Department-wise Activities"
          ? "Total activity volume by department"
          : onlyItem?.title === "Activity Type Distribution"
            ? "Combined activity categories"
            : "Activities by semester"
    );
  }
}

function drawPdfChartPage(doc, item, title = "Graph", startY = 88) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const cardWidth = pageWidth - marginX * 2;
  const footerTop = pageHeight - 20;
  const cardHeight = footerTop - startY;

  doc.setFillColor(255, 255, 255);
  doc.setDrawColor(193, 211, 203);
  doc.setLineWidth(0.5);
  doc.roundedRect(marginX, startY, cardWidth, cardHeight, 3, 3, "FD");

  doc.setFillColor(244, 248, 246);
  doc.roundedRect(marginX + 1.5, startY + 1.5, cardWidth - 3, 11, 2, 2, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12.5);
  doc.setTextColor(9, 31, 84);
  doc.text(item.title || "Graph", marginX + 5, startY + 8.4);

  addPdfImageContain(doc, item.image, marginX + 4, startY + 14, cardWidth - 8, cardHeight - 18);
}


async function generatePdfReport() {
  if (!window.jspdf || !window.jspdf.jsPDF) {
    alert("PDF library could not be loaded. Please check the internet connection and try again.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const sections = reportSections().filter(section => section.rows.length);
  const logoDataUrl = await getReportLogoDataUrl();

  if (!sections.length) {
    alert("There is no data in this report selection.");
    return;
  }

  addPdfCoverPage(doc, sections, logoDataUrl);

  for (let index = 0; index < sections.length; index++) {
    const section = sections[index];
    const meta = reportMeta(section.rows);
    const breakdown = activityBreakdown(section.rows);
    const facultyBreakdown = reportGroupBy(section.rows, "faculty");
    const departmentBreakdown = reportGroupBy(section.rows, "department");
    const focalPeople = focalPersonBreakdown(section.rows);
    const sectionTotal = reportAmount(section.rows);

    // A faculty report always begins with its own clean divider page.
    doc.addPage();
    if (section.kind === "faculty") {
      addFacultyDividerPage(doc, section, logoDataUrl);
      doc.addPage();
    }

    // ---------- SUMMARY PAGE ----------
    addPdfHeader(doc, section, index, sections.length, "ACTIVITY REPORT", logoDataUrl);
    const summaryStartY = 58;

    doc.autoTable({
      startY: summaryStartY,
      head: [["REPORT SUMMARY", "VALUE", "REPORT SUMMARY", "VALUE"]],
      body: reportMetaTableRows(meta),
      theme: "grid",
      styles: {
        fontSize: 9.4,
        fontStyle: "bold",
        cellPadding: 3.4,
        lineColor: [193, 204, 218],
        lineWidth: 0.25,
        textColor: [15, 23, 42]
      },
      headStyles: {
        fillColor: [9, 31, 84],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        fontSize: 9.8
      },
      alternateRowStyles: { fillColor: [248, 250, 252] },
      columnStyles: {
        0: { cellWidth: 83 },
        1: { cellWidth: 30, halign: "right" },
        2: { cellWidth: 83 },
        3: { cellWidth: 30, halign: "right" }
      },
      margin: { left: 14, right: 14 }
    });

    // ---------- FACULTY-WISE BREAKDOWN PAGE ----------
    if (facultyBreakdown.length) {
      doc.addPage();
      const facultyTableSection = { ...section, subtitle: `${section.subtitle || ""} - Faculty-wise Activities` };
      addPdfHeader(doc, facultyTableSection, index, sections.length, "FACULTY-WISE ACTIVITIES", logoDataUrl);
      const facultyStartY = 58;
      doc.autoTable({
        startY: facultyStartY,
        head: [["FACULTY", "TOTAL ACTIVITIES", "PERCENTAGE"]],
        body: facultyBreakdown.map(([faculty, total]) => [
          faculty,
          Number(total).toLocaleString(),
          reportPercentage(total, sectionTotal)
        ]),
        theme: "grid",
        styles: {
          fontSize: 9.1,
          fontStyle: "bold",
          cellPadding: 3.2,
          lineColor: [193, 204, 218],
          lineWidth: 0.22,
          overflow: "linebreak",
          textColor: [15, 23, 42]
        },
        headStyles: {
          fillColor: [9, 31, 84],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 9.5
        },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        columnStyles: {
          0: { cellWidth: 176 },
          1: { cellWidth: 38, halign: "right", fontStyle: "bold" },
          2: { cellWidth: 38, halign: "right", fontStyle: "bold" }
        },
        margin: { left: 14, right: 14, top: facultyStartY, bottom: 18 },
        didDrawPage: () => {
          addPdfHeader(doc, facultyTableSection, index, sections.length, "FACULTY-WISE ACTIVITIES", logoDataUrl);
        }
      });
    }

    // ---------- ACTIVITY TYPE BREAKDOWN PAGE ----------
    if (breakdown.length) {
      doc.addPage();
      const typeSection = { ...section, subtitle: `${section.subtitle || ""} - Activity Type Breakdown` };
      addPdfHeader(doc, typeSection, index, sections.length, "ACTIVITY BREAKDOWN", logoDataUrl);
      const typeStartY = 58;
      doc.autoTable({
        startY: typeStartY,
        head: [["ACTIVITY TYPE", "TOTAL", "PERCENTAGE"]],
        body: breakdown.map(([type, total]) => [
          type,
          total.toLocaleString(),
          reportPercentage(total, sectionTotal)
        ]),
        theme: "grid",
        styles: {
          fontSize: 9.1,
          fontStyle: "bold",
          cellPadding: 3.2,
          lineColor: [193, 204, 218],
          lineWidth: 0.22,
          overflow: "linebreak",
          textColor: [15, 23, 42]
        },
        headStyles: {
          fillColor: [11, 107, 58],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 9.5
        },
        alternateRowStyles: { fillColor: [246, 251, 248] },
        columnStyles: {
          0: { cellWidth: 176 },
          1: { cellWidth: 38, halign: "right", fontStyle: "bold" },
          2: { cellWidth: 38, halign: "right", fontStyle: "bold" }
        },
        margin: { left: 14, right: 14, top: typeStartY, bottom: 18 },
        didDrawPage: () => {
          addPdfHeader(doc, typeSection, index, sections.length, "ACTIVITY BREAKDOWN", logoDataUrl);
        }
      });
    }

    // ---------- DEPARTMENT-WISE BREAKDOWN PAGE ----------
    if (departmentBreakdown.length) {
      doc.addPage();
      const departmentTableSection = { ...section, subtitle: `${section.subtitle || ""} - Department-wise Activities` };
      addPdfHeader(doc, departmentTableSection, index, sections.length, "DEPARTMENT-WISE ACTIVITIES", logoDataUrl);
      const departmentStartY = 58;
      doc.autoTable({
        startY: departmentStartY,
        head: [["DEPARTMENT", "TOTAL ACTIVITIES", "PERCENTAGE"]],
        body: departmentBreakdown.map(([department, total]) => [
          department,
          Number(total).toLocaleString(),
          reportPercentage(total, sectionTotal)
        ]),
        theme: "grid",
        styles: {
          fontSize: 9.0,
          fontStyle: "bold",
          cellPadding: 3.0,
          lineColor: [193, 204, 218],
          lineWidth: 0.22,
          overflow: "linebreak",
          textColor: [15, 23, 42]
        },
        headStyles: {
          fillColor: [9, 31, 84],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 9.4
        },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        columnStyles: {
          0: { cellWidth: 176 },
          1: { cellWidth: 38, halign: "right", fontStyle: "bold" },
          2: { cellWidth: 38, halign: "right", fontStyle: "bold" }
        },
        margin: { left: 14, right: 14, top: departmentStartY, bottom: 18 },
        didDrawPage: () => {
          addPdfHeader(doc, departmentTableSection, index, sections.length, "DEPARTMENT-WISE ACTIVITIES", logoDataUrl);
        }
      });
    }

    // ---------- EXPECTED OUTCOMES ----------
    doc.addPage();
    addPdfExpectedOutcomesPage(doc, section, index, sections.length, logoDataUrl);

    // ---------- REMARKS ----------
    doc.addPage();
    addPdfRemarksPage(doc, section, index, sections.length, logoDataUrl);

    // ---------- GRAPH PAGES: ONE GRAPH PER PAGE FOR MAXIMUM VISIBILITY ----------
    const sectionCharts = await buildSectionCharts(section);

    const orderedCharts = [
      sectionCharts.find(item => item.title === "Faculty-wise Activities"),
      sectionCharts.find(item => item.title === "Activity Type Distribution"),
      sectionCharts.find(item => item.title === "Department-wise Activities"),
      sectionCharts.find(item => item.title === "Semester Trend")
    ].filter(Boolean);

    orderedCharts.forEach((chartItem, chartIndex) => {
      doc.addPage();
      const graphSection = {
        ...section,
        subtitle: `Graph page ${chartIndex + 1} of ${orderedCharts.length}`
      };
      addPdfHeader(doc, graphSection, index, sections.length, "GRAPHICAL SUMMARY", logoDataUrl);
      const graphStartY = 56;
      drawPdfChartPage(doc, chartItem, chartItem.title || "Graph", graphStartY);
    });

    // ---------- DETAILED RECORDS ----------
    doc.addPage();
    const detailSection = { ...section, subtitle: `${section.subtitle || ""} - Detailed Records` };
    addPdfHeader(doc, detailSection, index, sections.length, "DETAILED RECORDS", logoDataUrl);
    const detailsStartY = 58;

    doc.autoTable({
      startY: detailsStartY,
      head: [["SEMESTER", "FACULTY", "DEPARTMENT", "ACTIVITY TYPE", "HOW MANY", "ORGANIZATION", "FOCAL PERSON", "DESIGNATION", "SCHEDULED"]],
      body: detailRows(section.rows),
      theme: "grid",
      styles: {
        fontSize: 7.9,
        cellPadding: 2.6,
        overflow: "linebreak",
        lineColor: [190, 202, 216],
        lineWidth: 0.22,
        textColor: [15, 23, 42]
      },
      headStyles: {
        fillColor: [9, 31, 84],
        fontSize: 8.2,
        fontStyle: "bold",
        textColor: [255, 255, 255]
      },
      alternateRowStyles: { fillColor: [248, 250, 252] },
      columnStyles: {
        0: { cellWidth: 23, fontStyle: "bold" },
        1: { cellWidth: 34 },
        2: { cellWidth: 36 },
        3: { cellWidth: 29, fontStyle: "bold" },
        4: { cellWidth: 15, halign: "right", fontStyle: "bold" },
        5: { cellWidth: 45 },
        6: { cellWidth: 29 },
        7: { cellWidth: 29 },
        8: { cellWidth: 17, halign: "center" }
      },
      margin: { left: 14, right: 14, top: detailsStartY, bottom: 18 },
      didDrawPage: () => {
        // Repeat the report identity and filter context on every detail continuation page.
        addPdfHeader(doc, detailSection, index, sections.length, "DETAILED RECORDS", logoDataUrl);
      }
    });
  }

  addPdfFooter(doc);
  doc.save(reportFileName("pdf"));
}

function buildPrintGraphsHtml(charts) {
  if (!charts.length) return "";

  return `
    <h2>Graphs</h2>
    <div class="graph-grid">
      ${charts.map(item => `
        <div class="graph-card">
          <div class="graph-title">${esc(item.title)}</div>
          <img src="${item.image}" alt="${esc(item.title)}">
        </div>
      `).join("")}
    </div>`;
}

function buildPrintFocalPersonsHtml(focalPeople) {
  return `
    <h2>Focal Person Summary</h2>
    <table class="focal-table"><thead><tr><th>Focal Person</th><th>Designation</th><th>Department</th><th>Total</th></tr></thead><tbody>
      ${focalPeople.length
        ? focalPeople.map(item => `<tr><td>${esc(item.name)}</td><td>${esc(item.designation)}</td><td>${esc(item.department)}</td><td>${esc(item.total.toLocaleString())}</td></tr>`).join("")
        : `<tr><td>-</td><td>-</td><td>-</td><td>0</td></tr>`}
    </tbody></table>`;
}

async function printReport() {
  const sections = reportSections().filter(section => section.rows.length);
  if (!sections.length) {
    alert("There is no data in this report selection.");
    return;
  }

  const popup = window.open("", "_blank");
  if (!popup) {
    alert("Please allow pop-ups for this page to print the report.");
    return;
  }

  const logoDataUrl = await getReportLogoDataUrl();
  const sectionHtmlParts = [];

  for (const section of sections) {
    const meta = reportMeta(section.rows);
    const breakdown = activityBreakdown(section.rows);
    const facultyBreakdown = reportGroupBy(section.rows, "faculty");
    const departmentBreakdown = reportGroupBy(section.rows, "department");
    const sectionTotal = reportAmount(section.rows);
    const rows = detailRows(section.rows);
    const charts = await buildSectionCharts(section);

    sectionHtmlParts.push(`
      <section class="report-section">
        <div class="report-head">
          <div class="head-left">
            ${logoDataUrl ? `<img class="report-logo" src="${logoDataUrl}" alt="UOL Logo">` : ""}
            <div>
              <div class="university">THE UNIVERSITY OF LAHORE</div>
              <div class="office">External Linkages Intelligence Management System</div>
            </div>
          </div>
          <div class="report-label">ACTIVITY REPORT</div>
        </div>
        <h1>${esc(section.title)}</h1>
        <div class="subtitle">${esc(section.subtitle || "")}</div>
        <div class="generated-on">Generated on: ${esc(reportGeneratedOn())}</div>

        <div class="print-filters">
          ${reportFilterDetails(section).map(([label,value]) => `<div class="print-filter"><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`).join("")}
        </div>

        <div class="kpis">
          ${meta.map(item => `<div class="kpi"><span>${esc(item[0])}</span><strong>${esc(item[1])}</strong></div>`).join("")}
        </div>

        <div class="report-subsection page-section">
          <h2>Faculty-wise Activities</h2>
          <table class="small faculty-table"><thead><tr><th>Faculty</th><th>Total Activities</th><th>Percentage</th></tr></thead><tbody>
            ${facultyBreakdown.length
              ? facultyBreakdown.map(([faculty,total]) => `<tr><td>${esc(faculty)}</td><td>${esc(Number(total).toLocaleString())}</td><td>${esc(reportPercentage(total, sectionTotal))}</td></tr>`).join("")
              : `<tr><td>-</td><td>0</td><td>0.0%</td></tr>`}
          </tbody></table>
        </div>

        <div class="report-subsection page-section">
          <h2>Activity Breakdown</h2>
          <table class="small activity-table"><thead><tr><th>Activity Type</th><th>Total</th><th>Percentage</th></tr></thead><tbody>
            ${breakdown.length
              ? breakdown.map(([type,total]) => `<tr><td>${esc(type)}</td><td>${esc(Number(total).toLocaleString())}</td><td>${esc(reportPercentage(total, sectionTotal))}</td></tr>`).join("")
              : `<tr><td>-</td><td>0</td><td>0.0%</td></tr>`}
          </tbody></table>
        </div>

        <div class="report-subsection page-section">
          <h2>Department-wise Activities</h2>
          <table class="small department-table"><thead><tr><th>Department</th><th>Total Activities</th><th>Percentage</th></tr></thead><tbody>
            ${departmentBreakdown.length
              ? departmentBreakdown.map(([department,total]) => `<tr><td>${esc(department)}</td><td>${esc(Number(total).toLocaleString())}</td><td>${esc(reportPercentage(total, sectionTotal))}</td></tr>`).join("")
              : `<tr><td>-</td><td>0</td><td>0.0%</td></tr>`}
          </tbody></table>
        </div>

        ${buildPrintGraphsHtml(charts)}

        <div class="page-section detailed-records-section">
        <h2>Detailed Records</h2>
        <table><thead><tr><th>Semester</th><th>Faculty</th><th>Department</th><th>Activity Type</th><th>How Many</th><th>Organization</th><th>Focal Person</th><th>Designation</th><th>Scheduled</th></tr></thead><tbody>
          ${rows.map(cols => `<tr>${cols.map(col => `<td>${esc(col)}</td>`).join("")}</tr>`).join("")}
        </tbody></table>
        </div>
      </section>`);
  }

  popup.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(reportFileName("pdf").replace(".pdf", ""))}</title>
    <style>
      @page{size:A4 landscape;margin:12mm}
      *{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#000000;margin:0;background:#fff}
      .report-section{page-break-after:always}.report-section:last-child{page-break-after:auto}
      .report-head{display:flex;justify-content:space-between;align-items:center;background:linear-gradient(180deg,#091f54,#0b2d74);color:white;padding:14px 16px;border-radius:10px;border-bottom:5px solid #0b6b3a}
      .head-left{display:flex;align-items:center;gap:12px}.report-logo{width:56px;height:56px;border-radius:50%;background:#fff;padding:2px}
      .university{font-size:18px;font-weight:800;letter-spacing:.4px}.office{font-size:11px;margin-top:4px}.report-label{font-size:12px;font-weight:800;background:#fff;color:#091f54;padding:8px 12px;border-radius:20px}
      h1{font-size:20px;margin:18px 0 2px}.subtitle{color:#000000;font-size:11px;margin-bottom:4px}.generated-on{font-size:10px;color:#000000;margin-bottom:12px}h2{font-size:13px;margin:16px 0 7px}
      .print-filters{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin:10px 0 12px}.print-filter{border:1px solid #c1d3cb;border-radius:7px;padding:7px;background:#f8fafc;min-height:48px}.print-filter span{display:block;font-size:8px;font-weight:700;color:#0b6b3a;text-transform:uppercase;margin-bottom:4px}.print-filter strong{display:block;font-size:9px;color:#000;line-height:1.25}.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:7px}.kpi{border:1px solid #dbe4f0;border-radius:7px;padding:8px;background:#f8fafc}.kpi span{display:block;font-size:9px;color:#000000}.kpi strong{font-size:15px}
      .summary-grid{display:grid;grid-template-columns:1fr 1.1fr;gap:12px;align-items:start}
      .report-subsection{margin-top:14px}.page-section{break-before:page;page-break-before:always}.report-subsection:first-of-type{break-before:auto;page-break-before:auto}
      .faculty-table th,.department-table th{background:#091f54}.activity-table th{background:#0b6b3a}
      .graph-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin:10px 0 14px}.graph-card{border:1px solid #dbe4f0;border-radius:8px;padding:8px;background:#fff}.graph-title{font-size:11px;font-weight:700;margin-bottom:6px}.graph-card img{width:100%;height:250px;object-fit:contain;display:block;background:#fff}
      table{width:100%;border-collapse:collapse;font-size:8px;table-layout:fixed}th{background:#17365d;color:white;text-align:left;padding:5px;border:1px solid #dbe4f0}td{padding:4px;border:1px solid #dbe4f0;vertical-align:top;word-break:break-word}.small,.focal-table{width:100%;table-layout:auto}.small th{background:#7c3aed}.focal-table th{background:#0b6b3a}
      @media print{button{display:none}}
    </style></head><body>${sectionHtmlParts.join("")}<script>window.onload=()=>{window.print();};<\/script></body></html>`);
  popup.document.close();
}
// Bind report buttons after the original dashboard DOM has been initialized.
window.addEventListener("DOMContentLoaded", () => {
  $("pdfBtn")?.addEventListener("click", generatePdfReport);
});
