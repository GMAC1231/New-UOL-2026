/* =========================================================
   UOL TENTATIVE DASHBOARD — FULL APP.JS

   ACTIVITY WORKBOOK:
   data/UOL_Separated_All_Activity_Data_Cleaned.xlsx

   INTERNSHIP / PLACEMENT BREAKDOWN:
   Built directly from the main cleaned workbook.

   FILTERS:
   Faculty -> Department -> Activity Type -> Reporting Month

   Internship / Placement KPI:
   - Built from the main cleaned workbook
   - Changes by selected Faculty
   - Changes by selected Department
   - Does NOT show zero-value categories
   ========================================================= */

let allRows = [];
let filteredRows = [];

let allIpRows = [];
let ipWorkbookLoaded = false;

const charts = {};
const $ = id => document.getElementById(id);


/* =========================================================
   HELPERS
   ========================================================= */

function clean(value) {
  if (value === undefined || value === null) return "";

  const text = String(value).trim();

  if (
    !text ||
    ["undefined", "null", "nan"].includes(text.toLowerCase())
  ) {
    return "";
  }

  return text;
}

function norm(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function val(row, names) {
  const keys = Object.keys(row);

  for (const name of names) {
    const key = keys.find(k => norm(k) === norm(name));

    if (key !== undefined && clean(row[key])) {
      return clean(row[key]);
    }
  }

  return "";
}

function num(value) {
  const match = clean(value)
    .replace(/,/g, "")
    .match(/-?\d+(?:\.\d+)?/);

  return match ? Number(match[0]) : 0;
}

function yes(value) {
  return ["yes", "y", "true", "1", "scheduled"]
    .includes(clean(value).toLowerCase());
}

function esc(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function isInvalidLabel(value) {
  const text = norm(value);

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
  const type = norm(value);

  if (type === "glit" || type.includes("guest lecture")) {
    return "Guest Lecture / GLIT";
  }

  if (type === "iv" || type.includes("industrial visit")) {
    return "Industrial Visit / IV";
  }

  if (
    type === "ip" ||
    type === "internship placement" ||
    type.includes("internship") ||
    type.includes("intership") ||
    type.includes("placement")
  ) {
    return "Internship / Placement";
  }

  if (
    type === "mou" ||
    type.includes("mou signing") ||
    type.includes("memorandum of understanding")
  ) {
    return "MoU / MoU Signing";
  }

  if (
    type.includes("iab") ||
    type.includes("industry consultation")
  ) {
    return "IAB / Industry Consultation";
  }

  if (type.includes("research collaboration")) {
    return "Research Collaboration";
  }

  if (type.includes("community engagement")) {
    return "Community Engagement";
  }

  if (type.includes("alumni")) {
    return "Alumni Talk / Mentoring";
  }

  if (type.includes("curriculum feedback")) {
    return "Curriculum Feedback";
  }

  return clean(value);
}


/* =========================================================
   PARSE MAIN ACTIVITY DATA
   ========================================================= */

function parseActivities(rows) {
  const parsed = rows.map(row => {
    const rawType = val(row, [
      "Separated Activity Type",
      "Activity Type"
    ]);

    return {
      raw: row,

      ref: val(row, [
        "Ref #",
        "Ref",
        "Reference"
      ]),

      month: val(row, [
        "Reporting Month",
        "Month"
      ]),

      faculty: val(row, ["Faculty"]),

      department: val(row, ["Department"]),

      person: val(row, [
        "Focal Person",
        "Submitted By",
        "Person",
        "Name"
      ]),

      designation: val(row, ["Designation"]),

      originalType: rawType,

      type: standardizeActivityType(rawType),

      events: num(
        val(row, [
          "How Many",
          "Planned Quantity",
          "Total Activities"
        ])
      ),

      organization: val(row, [
        "Organization / Activity Title",
        "Organization / Activity",
        "Name Of Organization Planned with",
        "Name of Organization Planned with",
        "Planned Partner / Organization",
        "Organization"
      ]),

      planned: val(row, [
        "Planned",
        "Planned?"
      ]),

      scheduled: val(row, [
        "Scheduled",
        "Scheduled?"
      ]),

      remarks: val(row, [
        "Remarks / Result",
        "Remarks/Result",
        "Actual Outcome",
        "Remarks"
      ])
    };
  });

  return parsed.filter(row => (
    !isInvalidLabel(row.faculty) &&
    !isInvalidLabel(row.department) &&
    !isInvalidLabel(row.type)
  ));
}


/* =========================================================
   BUILD INTERNSHIP / PLACEMENT BREAKDOWN
   DIRECTLY FROM MAIN ACTIVITY DATA
   ========================================================= */

function buildIpBreakdownFromActivities() {
  const grouped = {};

  allRows.forEach(row => {

    if (
      row.type !== "Internship / Placement"
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
      norm(row.originalType) === "ip"
    ) {
      sourceType =
        "IP";
    }

    const key = [
      row.faculty,
      row.department,
      sourceType
    ].join("||");

    if (!grouped[key]) {
      grouped[key] = {
        faculty:
          row.faculty,

        department:
          row.department,

        sourceType:
          sourceType,

        students:
          0
      };
    }

    grouped[key].students +=
      row.events;
  });

  allIpRows =
    Object.values(
      grouped
    );

  ipWorkbookLoaded =
    true;

  const summary =
    summarizeIpRows(
      allIpRows
    );

  console.log(
    "Internship / Placement breakdown rows:",
    allIpRows.length
  );

  console.log(
    "Internship / Placement student total:",
    summary.total
  );
}


/* =========================================================
   UNIQUE / SELECT HELPERS
   ========================================================= */

function unique(values) {
  return [
    ...new Set(
      values
        .map(clean)
        .filter(
          value =>
            !isInvalidLabel(
              value
            )
        )
    )
  ].sort(
    (a, b) =>
      a.localeCompare(b)
  );
}


function fillSelect(
  id,
  values,
  label
) {

  const element =
    $(id);

  if (!element) {
    return;
  }

  element.innerHTML =
    `<option value="All">${label}</option>`;

  unique(values)
    .forEach(
      value => {

        const option =
          document.createElement(
            "option"
          );

        option.value =
          value;

        option.textContent =
          value;

        element.appendChild(
          option
        );
      }
    );
}


/* =========================================================
   CASCADING FILTER:
   DEPARTMENT
   ========================================================= */

function updateDepartmentFilter() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const select =
    $("departmentFilter");

  if (!select) {
    return;
  }

  const previous =
    select.value;

  let rows =
    [...allRows];

  if (
    faculty !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.faculty ===
          faculty
      );
  }

  const departments =
    unique(
      rows.map(
        row =>
          row.department
      )
    );

  select.innerHTML =
    `<option value="All">All Departments</option>`;

  departments.forEach(
    department => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        department;

      option.textContent =
        department;

      select.appendChild(
        option
      );
    }
  );

  select.value =
    (
      previous !== "All" &&
      departments.includes(
        previous
      )
    )
      ? previous
      : "All";
}


/* =========================================================
   CASCADING FILTER:
   ACTIVITY TYPE
   ========================================================= */

function updateActivityFilter() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const department =
    $("departmentFilter")
      ?.value ||
    "All";

  const select =
    $("activityFilter");

  if (!select) {
    return;
  }

  const previous =
    select.value;

  let rows =
    [...allRows];

  if (
    faculty !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.faculty ===
          faculty
      );
  }

  if (
    department !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.department ===
          department
      );
  }

  const activityTypes =
    unique(
      rows.map(
        row =>
          row.type
      )
    );

  select.innerHTML =
    `<option value="All">All Activity Types</option>`;

  activityTypes.forEach(
    activity => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        activity;

      option.textContent =
        activity;

      select.appendChild(
        option
      );
    }
  );

  select.value =
    (
      previous !== "All" &&
      activityTypes.includes(
        previous
      )
    )
      ? previous
      : "All";
}


/* =========================================================
   CASCADING FILTER:
   REPORTING MONTH
   ========================================================= */

function updateMonthFilter() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const department =
    $("departmentFilter")
      ?.value ||
    "All";

  const activity =
    $("activityFilter")
      ?.value ||
    "All";

  const select =
    $("monthFilter");

  if (!select) {
    return;
  }

  const previous =
    select.value;

  let rows =
    [...allRows];

  if (
    faculty !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.faculty ===
          faculty
      );
  }

  if (
    department !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.department ===
          department
      );
  }

  if (
    activity !== "All"
  ) {

    rows =
      rows.filter(
        row =>
          row.type ===
          activity
      );
  }

  const months =
    unique(
      rows
        .map(
          row =>
            row.month
        )
        .filter(
          month =>
            !isInvalidLabel(
              month
            )
        )
    );

  select.innerHTML =
    `<option value="All">All Reporting Months</option>`;

  months.forEach(
    month => {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        month;

      option.textContent =
        month;

      select.appendChild(
        option
      );
    }
  );

  select.value =
    (
      previous !== "All" &&
      months.includes(
        previous
      )
    )
      ? previous
      : "All";
}


/* =========================================================
   POPULATE FILTERS
   ========================================================= */

function populateFilters() {

  fillSelect(
    "facultyFilter",

    allRows.map(
      row =>
        row.faculty
    ),

    "All Faculties"
  );

  updateDepartmentFilter();

  updateActivityFilter();

  updateMonthFilter();
}


/* =========================================================
   APPLY FILTERS
   ========================================================= */

function applyFilters() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const department =
    $("departmentFilter")
      ?.value ||
    "All";

  const type =
    $("activityFilter")
      ?.value ||
    "All";

  const month =
    $("monthFilter")
      ?.value ||
    "All";

  const search =
    (
      $("searchFilter")
        ?.value ||
      ""
    )
      .toLowerCase()
      .trim();

  filteredRows =
    allRows.filter(
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

        const activityMatch =
          (
            type === "All" ||
            row.type ===
              type
          );

        const monthMatch =
          (
            month === "All" ||
            row.month ===
              month
          );

        const searchMatch =
          (
            !search ||

            [
              row.faculty,
              row.department,
              row.type,
              row.person,
              row.designation,
              row.organization,
              row.remarks,
              row.month
            ].some(
              value =>
                String(
                  value ||
                  ""
                )
                  .toLowerCase()
                  .includes(
                    search
                  )
            )
          );

        return (
          facultyMatch &&
          departmentMatch &&
          activityMatch &&
          monthMatch &&
          searchMatch
        );
      }
    );

  render();
}


/* =========================================================
   FILTER HANDLERS
   ========================================================= */

function handleFacultyChange() {

  updateDepartmentFilter();

  updateActivityFilter();

  updateMonthFilter();

  applyFilters();
}


function handleDepartmentChange() {

  updateActivityFilter();

  updateMonthFilter();

  applyFilters();
}


function handleActivityChange() {

  updateMonthFilter();

  applyFilters();
}


function handleMonthChange() {

  applyFilters();
}


/* =========================================================
   RESET FILTERS
   ========================================================= */

function resetFilters() {

  if (
    $("facultyFilter")
  ) {

    $("facultyFilter")
      .value =
      "All";
  }

  updateDepartmentFilter();

  if (
    $("departmentFilter")
  ) {

    $("departmentFilter")
      .value =
      "All";
  }

  updateActivityFilter();

  if (
    $("activityFilter")
  ) {

    $("activityFilter")
      .value =
      "All";
  }

  updateMonthFilter();

  if (
    $("monthFilter")
  ) {

    $("monthFilter")
      .value =
      "All";
  }

  if (
    $("searchFilter")
  ) {

    $("searchFilter")
      .value =
      "";
  }

  filteredRows =
    [...allRows];

  render();
}


/* =========================================================
   INTERNSHIP / PLACEMENT FILTERING
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

      return (
        facultyMatch &&
        departmentMatch
      );
    }
  );
}


/* =========================================================
   SUMMARIZE INTERNSHIP / PLACEMENT
   ========================================================= */

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


/* =========================================================
   HIDE / SHOW IMPACT ROW
   ========================================================= */

function setImpactRowVisible(
  id,
  visible
) {

  const el =
    $(id);

  if (
    !el
  ) {
    return;
  }

  const row =
    el.closest(
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


/* =========================================================
   CHANGE IMPACT LABEL
   ========================================================= */

function setImpactLabel(
  id,
  text
) {

  const el =
    $(id);

  if (
    !el
  ) {
    return;
  }

  const row =
    el.closest(
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


/* =========================================================
   RENDER INTERNSHIP / PLACEMENT SUMMARY
   ========================================================= */

function renderStudentSummary() {

  const ids = [

    "kInternshipStudents",

    "sideStudents",

    "impactStudents",

    "impactCombined",

    "impactIpRecords",

    "impactIpDepartments"

  ];

  if (
    !ipWorkbookLoaded
  ) {

    ids.forEach(
      id => {

        if (
          $(id)
        ) {

          $(id)
            .textContent =
            "—";
        }
      }
    );

    [
      "impactInternship",
      "impactPlacement",
      "impactPlanned"
    ].forEach(
      id =>
        setImpactRowVisible(
          id,
          false
        )
    );

    return;
  }

  const rows =
    getFilteredIpRows();

  const summary =
    summarizeIpRows(
      rows
    );

  if (
    $("kInternshipStudents")
  ) {

    $("kInternshipStudents")
      .textContent =
      summary.total
        .toLocaleString();
  }

  if (
    $("sideStudents")
  ) {

    $("sideStudents")
      .textContent =
      summary.total
        .toLocaleString();
  }

  if (
    $("impactStudents")
  ) {

    $("impactStudents")
      .textContent =
      summary.total
        .toLocaleString();
  }

  if (
    $("impactCombined")
  ) {

    $("impactCombined")
      .textContent =
      summary.total
        .toLocaleString();
  }

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

  const detailSlots = [

    "impactInternship",

    "impactPlacement",

    "impactPlanned"

  ];

  detailSlots.forEach(
    id => {

      setImpactRowVisible(
        id,
        false
      );
    }
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
      detailSlots.length
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
          detailSlots[
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
   GROUP DATA
   ========================================================= */

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

      const amount =
        row.events > 0
          ? row.events
          : 1;

      result[label] =
        (
          result[label] ||
          0
        ) +
        amount;
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


/* =========================================================
   DESTROY CHART
   ========================================================= */

function destroyChart(id) {

  if (
    charts[id]
  ) {

    charts[id]
      .destroy();

    delete charts[id];
  }
}


/* =========================================================
   WRAP LABEL
   ========================================================= */

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

  charts[id] =
    new Chart(
      canvas,
      {

        type:
          "bar",

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

                    return wrapLabel(
                      this
                        .getLabelForValue(
                          value
                        ),
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
   ACTIVITY DISTRIBUTION CHART
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

  const palette = [

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

  charts.activityChart =
    new Chart(
      canvas,
      {

        type:
          "doughnut",

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

                label(
                  context
                ) {

                  const value =
                    Number(
                      context.raw ||
                      0
                    );

                  const total =
                    context
                      .dataset
                      .data
                      .reduce(
                        (
                          sum,
                          item
                        ) =>
                          sum +
                          Number(
                            item ||
                            0
                          ),
                        0
                      );

                  const percent =
                    total > 0
                      ? (
                          (
                            value /
                            total
                          ) *
                          100
                        )
                          .toFixed(
                            1
                          )
                      : "0.0";

                  return (
                    `${context.label}: ` +
                    `${value.toLocaleString()} ` +
                    `(${percent}%)`
                  );
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


/* =========================================================
   ACTIVITY LEGEND
   ========================================================= */

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
      (
        sum,
        value
      ) =>
        sum +
        Number(
          value ||
          0
        ),
      0
    );

  legend.innerHTML =

    labels.map(
      (
        label,
        index
      ) => {

        const value =
          Number(
            values[index] ||
            0
          );

        const percent =
          total > 0
            ? (
                (
                  value /
                  total
                ) *
                100
              )
                .toFixed(
                  1
                )
            : "0.0";

        return `

          <button
            type="button"
            class="activity-legend-item"
            data-index="${index}"
            title="${esc(label)}: ${value.toLocaleString()} (${percent}%)"
          >

            <span
              class="activity-legend-dot"
              style="background:${colors[index]}"
            ></span>

            <span
              class="activity-legend-name"
            >
              ${esc(label)}
            </span>

            <strong
              class="activity-legend-value"
            >
              ${value.toLocaleString()}
            </strong>

          </button>

        `;
      }
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
              Number.isInteger(
                index
              ) &&
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
   MONTH HELPER
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


/* =========================================================
   MONTHLY TREND
   ========================================================= */

function monthlyChart() {

  const canvas =
    $("monthChart");

  if (
    !canvas
  ) {
    return;
  }

  destroyChart(
    "monthChart"
  );

  const result =
    {};

  filteredRows.forEach(
    row => {

      const month =
        monthKey(
          row.month
        );

      if (
        !month ||
        isInvalidLabel(
          month
        )
      ) {
        return;
      }

      const amount =
        row.events > 0
          ? row.events
          : 1;

      result[month] =
        (
          result[month] ||
          0
        ) +
        amount;
    }
  );

  const labels =
    Object.keys(
      result
    )
      .sort();

  charts.monthChart =
    new Chart(
      canvas,
      {

        type:
          "line",

        data: {

          labels:
            labels,

          datasets: [

            {

              data:

                labels.map(
                  month =>
                    result[
                      month
                    ]
                ),

              borderColor:
                "#0b6b3a",

              backgroundColor:
                "rgba(11,107,58,.10)",

              fill:
                true,

              tension:
                0.3,

              pointRadius:
                3
            }
          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          plugins: {

            legend: {

              display:
                false
            }
          },

          scales: {

            y: {

              beginAtZero:
                true,

              ticks: {

                precision:
                  0
              }
            },

            x: {

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
   RENDER DASHBOARD
   ========================================================= */

function render() {

  const totalActivities =
    filteredRows.reduce(
      (
        sum,
        row
      ) =>
        sum +
        (
          row.events > 0
            ? row.events
            : 1
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
          (
            row.events > 0
              ? row.events
              : 1
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
          (
            row.events > 0
              ? row.events
              : 1
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

  if (
    $("kRecords")
  ) {

    $("kRecords")
      .textContent =
      filteredRows
        .length
        .toLocaleString();
  }

  if (
    $("kActivities")
  ) {

    $("kActivities")
      .textContent =
      totalActivities
        .toLocaleString();
  }

  if (
    $("kScheduled")
  ) {

    $("kScheduled")
      .textContent =
      scheduled
        .toLocaleString();
  }

  if (
    $("kFaculties")
  ) {

    $("kFaculties")
      .textContent =
      faculties
        .size
        .toLocaleString();
  }

  if (
    $("kDepartments")
  ) {

    $("kDepartments")
      .textContent =
      departments
        .size
        .toLocaleString();
  }

  if (
    $("kTypes")
  ) {

    $("kTypes")
      .textContent =
      types
        .size
        .toLocaleString();
  }

  if (
    $("sideActivities")
  ) {

    $("sideActivities")
      .textContent =
      totalActivities
        .toLocaleString();
  }

  if (
    $("sideRecords")
  ) {

    $("sideRecords")
      .textContent =
      filteredRows
        .length
        .toLocaleString();
  }

  if (
    $("flowReported")
  ) {

    $("flowReported")
      .textContent =
      totalActivities
        .toLocaleString();
  }

  if (
    $("flowPlanned")
  ) {

    $("flowPlanned")
      .textContent =
      planned
        .toLocaleString();
  }

  if (
    $("flowScheduled")
  ) {

    $("flowScheduled")
      .textContent =
      scheduled
        .toLocaleString();
  }

  /* Internship / Placement */
  renderStudentSummary();

  /* Charts */
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
   FIND ACTIVITY SHEET
   ========================================================= */

function findActivitySheet(
  workbook
) {

  const preferredNames = [

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


/* =========================================================
   FETCH WORKBOOK
   ========================================================= */

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

    }

    catch(
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


/* =========================================================
   LOAD ACTIVITY WORKBOOK
   ========================================================= */

async function loadActivityWorkbook() {

  const paths = [

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

  filteredRows =
    [...allRows];

  /* Build IP breakdown from same Excel */
  buildIpBreakdownFromActivities();

  console.log(
    "Activity workbook:",
    result.path
  );

  console.log(
    "Activity sheet:",
    activitySheet
  );

  console.log(
    "Activity records:",
    allRows.length
  );

  return result.path;
}


/* =========================================================
   MASTER LOAD
   ========================================================= */

async function loadExcel() {

  if (
    $("connection")
  ) {

    $("connection")
      .innerHTML = `

        <span
          style="color:#f59e0b"
        >
          ●
        </span>

        Loading workbook…

      `;
  }

  try {

    /* Only ONE Excel file */
    await loadActivityWorkbook();

    populateFilters();

    render();

    const totalActivities =
      allRows.reduce(
        (
          sum,
          row
        ) =>
          sum +
          (
            row.events > 0
              ? row.events
              : 1
          ),
        0
      );

    const faculties =
      new Set(
        allRows
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
        allRows
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

    const ipSummary =
      summarizeIpRows(
        allIpRows
      );

    if (
      $("connection")
    ) {

      $("connection")
        .innerHTML = `

          <span
            style="color:#16a34a"
          >
            ●
          </span>

          Connected

          •

          ${allRows.length.toLocaleString()}
          records

          •

          ${totalActivities.toLocaleString()}
          activities

          <br>

          ${faculties.size.toLocaleString()}
          faculties

          •

          ${departments.size.toLocaleString()}
          departments

          •

          <strong>
            ${ipSummary.total.toLocaleString()}
            Internship / Placement students
          </strong>

        `;
    }
  }

  catch(
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
        .innerHTML = `

          <span
            style="color:#dc2626"
          >
            ●
          </span>

          Excel workbook failed

        `;
    }
  }
}


/* =========================================================
   POPUP DISPLAY VALUE
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


/* =========================================================
   DETAIL POPUP
   ========================================================= */

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

          <div
            class="detail-record"
          >

            <strong>
              Record ${index + 1}
            </strong>

            <div
              class="detail-grid"
            >

              <div class="detail-field">
                <small>
                  Faculty
                </small>
                <strong>
                  ${displayValue(
                    row.faculty
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Department
                </small>
                <strong>
                  ${displayValue(
                    row.department
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Activity Type
                </small>
                <strong>
                  ${displayValue(
                    row.type
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Original Activity Type
                </small>
                <strong>
                  ${displayValue(
                    row.originalType
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  How Many
                </small>
                <strong>
                  ${
                    Number(
                      row.events ||
                      0
                    )
                      .toLocaleString()
                  }
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Reporting Month
                </small>
                <strong>
                  ${displayValue(
                    row.month
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Person
                </small>
                <strong>
                  ${displayValue(
                    row.person
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Designation
                </small>
                <strong>
                  ${displayValue(
                    row.designation
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Organization
                </small>
                <strong>
                  ${displayValue(
                    row.organization
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Planned
                </small>
                <strong>
                  ${displayValue(
                    row.planned
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Scheduled
                </small>
                <strong>
                  ${displayValue(
                    row.scheduled
                  )}
                </strong>
              </div>

              <div class="detail-field">
                <small>
                  Remarks / Result
                </small>
                <strong>
                  ${displayValue(
                    row.remarks
                  )}
                </strong>
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


/* =========================================================
   CLOSE POPUP
   ========================================================= */

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


  /* Internship / Placement export */

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


  const ipSummary =
    summarizeIpRows(
      ipRows
    );

  const summaryExport =
    Object.entries(
      ipSummary.byType
    )
      .filter(
        (
          [
            ,
            students
          ]
        ) =>
          students > 0
      )
      .map(
        (
          [
            type,
            students
          ]
        ) => ({

          Category:
            type,

          Students:
            students
        })
      );

  summaryExport.push(
    {

      Category:
        "Combined Total",

      Students:
        ipSummary.total
    }
  );

  XLSX.utils
    .book_append_sheet(

      workbook,

      XLSX.utils
        .json_to_sheet(
          summaryExport
        ),

      "IP Summary"
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
   START DASHBOARD
   ========================================================= */

document
  .addEventListener(
    "DOMContentLoaded",
    () => {

      bindEvents();

      loadExcel();
    }
  );