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
                "Month"
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
      )
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

      const key =
        [
          row.faculty,
          row.department,
          sourceType
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
        .map(clean)
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

  select.value =
    (
      previous !== "All" &&
      items.includes(
        previous
      )
    )
      ? previous
      : "All";
}


/* =========================================================
   CASCADING FILTERS
   ========================================================= */

function updateFacultyFilter() {

  setSelectOptions(
    "facultyFilter",
    allRows.map(
      row =>
        row.faculty
    ),
    "All Faculties"
  );
}


function updateDepartmentFilter() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

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

  setSelectOptions(
    "departmentFilter",
    rows.map(
      row =>
        row.department
    ),
    "All Departments"
  );
}


function updateActivityFilter() {

  const faculty =
    $("facultyFilter")
      ?.value ||
    "All";

  const department =
    $("departmentFilter")
      ?.value ||
    "All";

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

  setSelectOptions(
    "activityFilter",
    rows.map(
      row =>
        row.type
    ),
    "All Activity Types"
  );
}


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

  setSelectOptions(
    "monthFilter",
    rows.map(
      row =>
        row.month
    ),
    "All Reporting Months"
  );
}


function updateOrganizationFilter() {

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

  const month =
    $("monthFilter")
      ?.value ||
    "All";

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

  if (
    month !== "All"
  ) {
    rows =
      rows.filter(
        row =>
          row.month ===
          month
      );
  }

  setSelectOptions(
    "organizationFilter",
    rows.map(
      row =>
        row.organization
    ),
    "All Organizations"
  );
}


function populateFilters() {

  updateFacultyFilter();

  if (
    $("facultyFilter")
  ) {
    $("facultyFilter")
      .value =
      "All";
  }

  updateDepartmentFilter();

  updateActivityFilter();

  updateMonthFilter();

  updateOrganizationFilter();
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

  const activity =
    $("activityFilter")
      ?.value ||
    "All";

  const month =
    $("monthFilter")
      ?.value ||
    "All";

  const organization =
    $("organizationFilter")
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
            activity === "All" ||
            row.type ===
            activity
          );

        const monthMatch =
          (
            month === "All" ||
            row.month ===
            month
          );

        const organizationMatch =
          (
            organization === "All" ||
            row.organization ===
            organization
          );

        const searchMatch =
          (
            !search ||
            [
              row.faculty,
              row.department,
              row.type,
              row.originalType,
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
          organizationMatch &&
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

  updateOrganizationFilter();

  applyFilters();
}


function handleDepartmentChange() {

  updateActivityFilter();

  updateMonthFilter();

  updateOrganizationFilter();

  applyFilters();
}


function handleActivityChange() {

  updateMonthFilter();

  updateOrganizationFilter();

  applyFilters();
}


function handleMonthChange() {

  updateOrganizationFilter();

  applyFilters();
}


function handleOrganizationChange() {

  applyFilters();
}


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

  updateOrganizationFilter();

  if (
    $("organizationFilter")
  ) {
    $("organizationFilter")
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
            ${Number(values[index] || 0).toLocaleString()}
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

      result[month] =
        (
          result[month] ||
          0
        ) +
        rowAmount(
          row
        );
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

  setText(
    "kRecords",
    filteredRows
      .length
      .toLocaleString()
  );

  setText(
    "kActivities",
    totalActivities
      .toLocaleString()
  );

  setText(
    "kScheduled",
    scheduled
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
    "sideActivities",
    totalActivities
      .toLocaleString()
  );

  setText(
    "sideRecords",
    filteredRows
      .length
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

  setText(
    "flowScheduled",
    scheduled
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