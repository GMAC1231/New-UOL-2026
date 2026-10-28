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
let pr41OrgEvidence = {organizations:[],reportedStudents:57};
let pr46OrgEvidence = {organizations:[],reportedStudents:102,evidenceStudents:98,namedCompanyRows:91};
const pr46OrgNames = () => pr46OrgEvidence.organizations.map(x=>x.name);
const isPr46Internship = row => row.ref === 'PRPSTFPT46' && /internship/i.test(String(row.type || '')) && (/^multiple$/i.test(String(row.organization || '').trim()) || /internship company names in evidence/i.test(String(row.organization || '')));
const pr41OrgNames = () => pr41OrgEvidence.organizations.map(x=>x.name);
const isPr41Internship = row => row.ref === 'PRPSTFPT41' && row.type === 'Internship / Placement' && /mentioned in excel sheet/i.test(String(row.organization || ''));
const executiveOrgAliases = row => isPr41Internship(row) ? pr41OrgNames() : isPr46Internship(row) ? pr46OrgNames() : [row.organization];

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

  if (type === "remote") return "Internship / Placement";

  if (
    type === "planned" ||
    type.includes(
      "planned activity"
    )
  ) {
    return "Internship / Placement";
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

      const semester = semesterForActivity(row);

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


function setSelectOptions(id, values, allLabel, counts = null, total = null) {
  const select = $(id);
  if (!select) return;
  const previous = select.value;
  const items = unique(values);
  select.replaceChildren();
  const makeOption = (value, label) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    select.appendChild(option);
  };
  const totalChoices = items.length.toLocaleString('en-US');
  makeOption('All', `${allLabel} (${totalChoices})${Number.isFinite(total) ? ' · '+Number(total).toLocaleString('en-US')+' activities' : ''}`);
  items.forEach(item => {
    const count = counts instanceof Map ? counts.get(item) : counts?.[item];
    const unfilteredInternship = id === "activityFilter" && item === "Internship / Placement" && isCompleteOverviewScope();
    makeOption(item, `${item} (${(unfilteredInternship ? CONFIRMED_IP_TOTAL : Number(count || 0)).toLocaleString('en-US')})`);
  });
  select.value = previous !== 'All' && items.includes(previous) ? previous : 'All';
}

function optionTotals(rows, key) {
  const totals = new Map();
  for (const row of rows) {
    const label = clean(row[key]);
    if (isInvalidLabel(label)) continue;
    totals.set(label, (totals.get(label) || 0) + rowAmount(row));
  }
  return totals;
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
    row => semesterForActivity(row) === semester
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
    [...rows.map(row => row.faculty), ...uolInstitutionalRoster.map(r => r.faculty)],
    "All Faculties", optionTotals(rows, "faculty"), reconciledTotal(rows, {wholeCohort:isCompleteOverviewScope()})
  );
}


/* =========================================================
   DEPARTMENT FILTER
   Depends on Semester + Faculty
   ========================================================= */

let uolInstitutionalRoster=[];
let uolSubmissionRoster=[];
async function loadInstitutionalRoster(){
 try{
  const response=await fetch('data/65_Submissions_Source.xlsx',{cache:'no-store'});
  if(!response.ok)throw Error('Submission workbook could not load');
  const book=XLSX.read(await response.arrayBuffer(),{type:'array'});
  const submissionRows=XLSX.utils.sheet_to_json(book.Sheets['All Submissions'],{defval:'',raw:false})
   .filter(r=>window.UOLSnapshots.contains(r['Ref #']));
  uolSubmissionRoster=submissionRows.map(r=>({ref:clean(r['Ref #']),faculty:clean(r.Faculty),department:clean(r.Department),month:clean(r['Reporting Month'])}));
  uolInstitutionalRoster=submissionRows
   .filter(r=>!isInvalidLabel(r.Department))
   .map(r=>({department:clean(r.Department),faculty:clean(r.Faculty)}));
 }catch(err){console.warn('Institutional department roster unavailable',err);}
}
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
    [...rows.map(row => row.department),...uolInstitutionalRoster.filter(r=>faculty==='All'||r.faculty===faculty).map(r=>r.department)],
    "All Departments", optionTotals(rows, "department"), reconciledTotal(rows, {wholeCohort:isCompleteOverviewScope()})
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
    // Keep all 13 named activity types available in Spring and Fall, even at zero.
    // Do not show the data-quality placeholder "Unclassified" as an activity type.
    allRows.map(row => row.type)
      .filter(type => !isInvalidLabel(type) && norm(type) !== "total activities"),
    "All Activity Types", optionTotals(rows.filter(r=>norm(r.type)!=="total activities"), "type"),
    reconciledTotal(rows.filter(r=>norm(r.type)!=="total activities"), {wholeCohort:isCompleteOverviewScope()})
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
  // Only 2026 semesters are available in this reporting dashboard.
  if (explicitYear && explicitYear !== "2026") return "";

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

  if (!month) return "";
  if (year && year !== "2026") return "";

  const semester = month <= 6 ? "Spring Semester" : "Fall Semester";
  return year ? `${semester} ${year}` : semester;
}

// Respect the original progress report semester for this July-dated Urdu submission.
// Its Reporting Month remains 2026-07-17; its Year field explicitly specifies Spring 2026.
// Provisional assignments from activity dates in the original 64-submission report.
// These do not alter blank Reporting Month fields or confirm completion of planned events.
const SEMESTER_FROM_ACTIVITY_EVIDENCE = Object.freeze({
  'PRPSTFPT50': 'Spring Semester 2026',
  'PRPSTFPT48': 'Spring Semester 2026',
  'PRPSTFPT47': 'Spring Semester 2026',
  'PRPSTFPT45': 'Spring Semester 2026',
  'PRPSTFPT39': 'Spring Semester 2026',
  'PRPSTFPT38': 'Spring Semester 2026',
  'PRPSTFPT37': 'Spring Semester 2026',
  'PRPSTFPT36': 'Spring Semester 2026',
  'PRPSTFPT35': 'Spring Semester 2026',
  'PRPSTFPT26': 'Spring Semester 2026',
  'PRPSTFPT103': 'Fall Semester 2026',
  'PRPSTFPT86': 'Fall Semester 2026',
  'PRPSTFPT85': 'Fall Semester 2026',
  'PRPSTFPT82': 'Fall Semester 2026',
});

function semesterForActivity(row) {
  if (typeof row === 'string') return semesterFromMonth(row);
  const ref = String(row?.ref || row?.['Ref #'] || '').trim();
  if (ref === 'PRPSTFPT34') return 'Spring Semester 2026';
  if (SEMESTER_FROM_ACTIVITY_EVIDENCE[ref]) return SEMESTER_FROM_ACTIVITY_EVIDENCE[ref];
  return semesterFromMonth(row?.month || row?.['Reporting Month']);
}

function updateMonthFilter() {

  setSelectOptions(
    "monthFilter",
    [...allRows.map(row => semesterForActivity(row)).filter(Boolean), "Spring Semester 2026", "Fall Semester 2026"],
    "All Semesters", (()=>{const map=new Map();for(const r of allRows){const k=semesterForActivity(r);if(k&&!isInvalidLabel(k))map.set(k,(map.get(k)||0)+rowAmount(r));}return map;})(), reconciledTotal(allRows)
  );
}


/* =========================================================
   ORGANIZATION FILTER
   Depends on Semester + Faculty + Department + Activity Type
   ========================================================= */

// The cleaned master worksheet is the sole authority for Executive Overview
// organization choices. Older PRPSTFPT41/46 evidence lists are not injected
// into this dropdown because their entries can be deleted or revised.
function currentOverviewOrganizations(row) {
  const raw = clean(row.organization);
  // Each PRPSTFPT41 organization now has its own canonical row and quantity.
  if (isInvalidLabel(raw) || /mentioned in excel sheet|about\s+sixty\s+six\s+companies\s+for\s+the\s+placement/i.test(raw)) return [];
  if (/^(?:unspecified|not specified|multiple|\.\.\.|n\/?a|none|-|—)$/i.test(raw)) return [];
  return [raw];
}

function updateOrganizationFilter() {
  const faculty = $("facultyFilter")?.value || "All";
  const department = $("departmentFilter")?.value || "All";
  const activity = $("activityFilter")?.value || "All";
  let rows = rowsForSemester(allRows);
  if (faculty !== "All") rows = rows.filter(row => row.faculty === faculty);
  if (department !== "All") rows = rows.filter(row => row.department === department);
  if (activity !== "All") rows = rows.filter(row => row.type === activity);
  const totals = organizationOptionTotals(rows);
  setSelectOptions("organizationFilter", [...totals.keys()], "All Organizations", totals,
    reconciledTotal(rows));
}

function organizationOptionTotals(rows) {
  const totals = new Map();
  for (const row of rows) {
    for (const name of currentOverviewOrganizations(row)) {
      totals.set(name, (totals.get(name) || 0) + rowAmount(row));
    }
  }
  return totals;
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
      month === "All" || semesterForActivity(row) === month;

    const organizationMatch =
      organization === "All" || currentOverviewOrganizations(row).includes(organization);

    const searchMatch =
      !search || [
        row.faculty,
        row.department,
        row.type,
        row.originalType,
        row.person,
        row.designation,
        row.organization,
        ...currentOverviewOrganizations(row),
        row.remarks,
        row.month,
        semesterForActivity(row)
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

  // Match Section 2's owner-confirmed submitted-list scope only when no
  // faculty, department, semester or organization restriction is active.
  // Filtered breakdowns still belong to the legacy cleaned worksheet.
  const executiveIpConfirmedScope =
    ($("facultyFilter")?.value || "All") === "All" &&
    ($("departmentFilter")?.value || "All") === "All" &&
    ($("monthFilter")?.value || "All") === "All" &&
    ($("organizationFilter")?.value || "All") === "All" &&
    !($('searchFilter')?.value || '').trim();
  const executiveIpVisibleTotal = executiveIpConfirmedScope ? CONFIRMED_IP_TOTAL : summary.total;

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
          executiveIpVisibleTotal
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

  const slots = ["impactInternship"];

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

// Canonical owner-confirmed quantity for the complete Internship / Placement scope.
// Only replace the *aggregate* when the view contains the entire source cohort.
// Do not distribute unattributed records across faculties, organizations or semesters.
// Original internship worksheet evidence, bundled for reliable offline PDF exports.
// Each record keeps its faculty, department, organization and submission reference.
const PDF_INTERNSHIP_SOURCE_ROWS = [{"ref":"PRPSTFPT77","submittedBy":"Foc Uidns Lhr","status":"Pending From Linkage Officer","submissionDate":"8 September 2026","reportingMonth":"2026-09-04","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"University Institute of Diet & Nutritional Sciences","person":"Foc Uidns Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"IP","events":1,"organization":"UOTH and Social Security Hospital Multan Road","planned":"Yes","scheduled":"Yes","remarks":"Internship completed"},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"","planned":"No","scheduled":"Yes","remarks":"document is attached at the end, How Many: 24"},{"ref":"PRPSTFPT46","submittedBy":"Foc Dell Lhr","status":"Pending From Linkage Officer","submissionDate":"22 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Languages & Literature","department":"Deptartment of English Language and Literature","person":"Foc Dell Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Eushbi Private Limited","planned":"Yes","scheduled":"Yes","remarks":"The internship at the Eushbi Private Limited is enhancing technical expertise and interpersonal skills through teamwork toward organizational goals and to understand environmental problems"},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"CEMB","planned":"Yes","scheduled":"Yes","remarks":"Developed technical and interpersonal skills while collaborating with team members to achieve organizational goals"},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":".Diagnostic Laboratories","planned":"Yes","scheduled":"Yes","remarks":"Gain practical industry experience and develop professional skills relevant to the chosen field"},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Jinnah hospitals .lAHORE","planned":"Yes","scheduled":"Yes","remarks":"Gain practical industry experience and develop professional skills relevant to the chosen field"},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"children’s hospital .Lahore","planned":"Yes","scheduled":"Yes","remarks":"Developed technical and interpersonal skills while collaborating with team members to achieve organizational goals"},{"ref":"PRPSTFPT40","submittedBy":"Foc Imbb Lhr","status":"Pending From Linkage Officer","submissionDate":"18 July 2026","reportingMonth":"2026-10-05","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Institute of Molecular Biology & Biotechnology","person":"Foc Imbb Lhr","designation":"Associate Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"PCSIR","planned":"Yes","scheduled":"Yes","remarks":"Students learn theoretical & Applied knowledge, technical / operational skills, soft skills & professional exposure, industry Awareness & Networking"},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"2 Space","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"AB Technologies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Ahmaar (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Async Digital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Async Digital Software House","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"AWS Tech","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Bytewise Technologies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Codexiom","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Corvit Systems","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Devsinc","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Dgware Solutions","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Ebryx (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"E-Care Solution","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Eyra Tech","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Fermi Soft","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Flicken","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"GulfBase","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Hoster99 (Web Hosting Company)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Ibex","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"IIFA Tech","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Infinity Wave","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Innosoft Works (SMC-Private) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Invosage","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Kashf Foundation","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Kings Solutions","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Nexrupt technologies pvt ltd","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"OMNI Soft","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Prismware Technologies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Qubit Dynamics","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Rstack Labs","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Strategic System","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Studio 93","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Tech Lynx Overdrive","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Thingtrax","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Tx Labs","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"University of Engineering & Technology (UET)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Upvave","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"US Mobile","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Wix Patriots","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Wordx BPO","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT39","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Z Space","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Azhar Engineering and Power Generation (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"M.S.B Engineering Services","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Machine Zone (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"NIMIR Group of Companies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Nishat Power Plant","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Project Cell (HVAC)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"S.T Engineering Services (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Saab Moulded Industry (MDK)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Toyota HN Motors","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT38","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Xida Mould","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Akhter Saeed Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Alam Medix Company","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Amarant Medical Technologies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Econ Whizz","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Farooq Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Jinnah Hospital Lahore","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"RM Scientific Lab","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Saleem Memorial Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Shaukat Khanum Memorial Cancer Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Shifa International Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT37","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"University of Lahore Teaching Hospital","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Al Riaz Civil Engineering Services (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Al Shabbir Developers","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"CCECC","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Eratech Engineering Company","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Lahore Engineering Consultant (LEC)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Luxurious Design & Built (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"MS Builder and Construction Co.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Nazim & Qasim (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT36","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Raza Engineers Construction","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT35","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Durable Technologies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT35","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Electros Energies","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT35","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Lahore Electric Supply Company (LESCO)","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT35","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Technology Links (Pvt.) Ltd.","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT35","submittedBy":"Foc Dot Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Technology","person":"Foc Dot Lhr","designation":"","type":"Internship / Placement","originalType":"IP","events":1,"organization":"Transfopower Industries","planned":"Yes","scheduled":"","remarks":""},{"ref":"PRPSTFPT34","submittedBy":"Foc Dou Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"2026-07-17","campus":"Lahore Campus","faculty":"Faculty of Languages & Literature","department":"Department of Urdu","person":"Foc Dou Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"1. Pakistan Academy of Letters","planned":"Yes","scheduled":"Yes","remarks":"Literary Research & Editorial Exposure"},{"ref":"PRPSTFPT30","submittedBy":"Foc Doed Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-16","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Education","person":"Foc Doed Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Govt. Girls High School Bhubtian","planned":"Yes","scheduled":"Yes","remarks":""},{"ref":"PRPSTFPT30","submittedBy":"Foc Doed Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-16","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Education","person":"Foc Doed Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Govt. High School Sher Shah Colony","planned":"Yes","scheduled":"Yes","remarks":""},{"ref":"PRPSTFPT30","submittedBy":"Foc Doed Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-16","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Education","person":"Foc Doed Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Govt. Girls Elementary School Sher Shah Colony","planned":"Yes","scheduled":"Yes","remarks":""},{"ref":"PRPSTFPT30","submittedBy":"Foc Doed Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-16","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Education","person":"Foc Doed Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Hamza Foundation","planned":"Yes","scheduled":"Yes","remarks":""},{"ref":"PRPSTFPT30","submittedBy":"Foc Doed Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-16","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Education","person":"Foc Doed Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"Child Welfare Center","planned":"Yes","scheduled":"Yes","remarks":""},{"ref":"PRPSTFPT101","submittedBy":"Foc Uimlt Lhr","status":"Approved","submissionDate":"5 October 2026","reportingMonth":"2026-09-21","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"University Institute of Medical Lab Technology","person":"Foc Uimlt Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT95","submittedBy":"Foc Dois Lhr","status":"Pending From Linkage Officer","submissionDate":"30 September 2026","reportingMonth":"2026-09-30","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Islamic Studies","person":"Fartashia Dastagir","designation":"Research Associate","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT95","submittedBy":"Foc Dois Lhr","status":"Pending From Linkage Officer","submissionDate":"30 September 2026","reportingMonth":"2026-09-30","campus":"Lahore Campus","faculty":"Faculty of Social Sciences","department":"Department of Islamic Studies","person":"Fartashia Dastagir","designation":"Research Associate","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT84","submittedBy":"Foc Lsn Lhr","status":"Pending From Linkage Officer","submissionDate":"15 September 2026","reportingMonth":"2026-02-02","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Lahore School of Nursing","person":"Foc Lsn Lhr","designation":"Assistant professor","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT79","submittedBy":"Foc Soca Lhr","status":"Pending From Linkage Officer","submissionDate":"10 September 2026","reportingMonth":"2026-09-10","campus":"Lahore Campus","faculty":"Faculty of Arts and Architecture","department":"School of Creative Arts","person":"Foc Sociology Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT76","submittedBy":"Foc Soca Lhr","status":"Pending From Linkage Officer","submissionDate":"17 August 2026","reportingMonth":"2026-05-21","campus":"Lahore Campus","faculty":"Faculty of Arts and Architecture","department":"School of Creative Arts","person":"Foc Soca Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT59","submittedBy":"Foc Lsa Lhr","status":"Pending From Linkage Officer","submissionDate":"25 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Management Sciences","department":"Lahore School of Aviation","person":"Foc Lsa Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT59","submittedBy":"Foc Lsa Lhr","status":"Pending From Linkage Officer","submissionDate":"25 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Management Sciences","department":"Lahore School of Aviation","person":"Foc Lsa Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT59","submittedBy":"Foc Lsa Lhr","status":"Pending From Linkage Officer","submissionDate":"25 July 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Management Sciences","department":"Lahore School of Aviation","person":"Foc Lsa Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT56","submittedBy":"Foc Dhpt Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Health Professional Technologies","person":"Foc Dhpt Lhr","designation":"Senior Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT55","submittedBy":"Foc Dovs Lhr","status":"Pending From Linkage Officer","submissionDate":"23 July 2026","reportingMonth":"2026-07-23","campus":"Lahore Campus","faculty":"Faculty of Allied Health Sciences","department":"Department of Optometry & Vision Science","person":"Foc Dovs Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT34","submittedBy":"Foc Dou Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"2026-07-17","campus":"Lahore Campus","faculty":"Faculty of Languages & Literature","department":"Department of Urdu","person":"Foc Dou Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT34","submittedBy":"Foc Dou Lhr","status":"Pending From Linkage Officer","submissionDate":"17 July 2026","reportingMonth":"2026-07-17","campus":"Lahore Campus","faculty":"Faculty of Languages & Literature","department":"Department of Urdu","person":"Foc Dou Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT28","submittedBy":"Foc Doms Lhr","status":"Pending From Linkage Officer","submissionDate":"16 July 2026","reportingMonth":"2026-07-14","campus":"Lahore Campus","faculty":"Faculty of Sciences","department":"Department of Mathematics and Statistics","person":"Foc Doms Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT21","submittedBy":"Foc Dce Lhr","status":"Pending From Linkage Officer","submissionDate":"9 July 2026","reportingMonth":"2026-07-01","campus":"Lahore Campus","faculty":"Faculty of Engineering & Technology","department":"Department of Civil Engineering","person":"Foc Dce Lhr","designation":"Assistant Professor","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT13","submittedBy":"Foc Law Lhr","status":"Pending From Linkage Officer","submissionDate":"9 July 2026","reportingMonth":"2026-07-09","campus":"Lahore Campus","faculty":"Faculty of Law","department":"M.A Raoof College of Law","person":"Foc Law Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT13","submittedBy":"Foc Law Lhr","status":"Pending From Linkage Officer","submissionDate":"9 July 2026","reportingMonth":"2026-07-09","campus":"Lahore Campus","faculty":"Faculty of Law","department":"M.A Raoof College of Law","person":"Foc Law Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT13","submittedBy":"Foc Law Lhr","status":"Pending From Linkage Officer","submissionDate":"9 July 2026","reportingMonth":"2026-07-09","campus":"Lahore Campus","faculty":"Faculty of Law","department":"M.A Raoof College of Law","person":"Foc Law Lhr","designation":"Lecturer","type":"Internship / Placement","originalType":"Internship","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT6","submittedBy":"Foc Lbs Lhr","status":"Rejected","submissionDate":"15 June 2026","reportingMonth":"2026-01-15","campus":"","faculty":"Faculty of Management Sciences","department":"Lahore Business School","person":"Foc Lbs Lhr","designation":"ASSISTAANT PROFESSOR","type":"Internship / Placement","originalType":"Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""},{"ref":"PRPSTFPT103","submittedBy":"Foc Soca Lhr","status":"Pending From Linkage Officer","submissionDate":"7 October 2026","reportingMonth":"","campus":"Lahore Campus","faculty":"Faculty of Arts and Architecture","department":"School of Creative Arts","person":"Foc Soca Lhr","designation":"","type":"Internship / Placement","originalType":"Internship/Placement","events":1,"organization":"","planned":"","scheduled":"","remarks":""}];
const PR41_PDF_ORGANIZATIONS = [{"name":"Code Alpha","evidenceRecords":4},{"name":"Texense","evidenceRecords":3},{"name":"Astute Software","evidenceRecords":2},{"name":"Bookme","evidenceRecords":2},{"name":"FlyRank AI (Remote Machine Learning Internship)","evidenceRecords":2},{"name":"NerithonX Technologies (PVT.) LTD.","evidenceRecords":2},{"name":"Shabbir Industries & Amir Steel","evidenceRecords":2},{"name":"Systems Limited Head Office (intern/trainee in Automations )","evidenceRecords":2},{"name":"10Pearls","evidenceRecords":1},{"name":"American Prime Solution","evidenceRecords":2},{"name":"Amrood labs","evidenceRecords":1},{"name":"Arch Technologies","evidenceRecords":1},{"name":"Arfa tower","evidenceRecords":1},{"name":"Arhamsoft","evidenceRecords":1},{"name":"CCI ( coca-cola Icecek)","evidenceRecords":1},{"name":"Claim Care","evidenceRecords":1},{"name":"Clicktake technologies","evidenceRecords":1},{"name":"Cobb Uk.","evidenceRecords":1},{"name":"Code Alpha (Remote) , Vision Plus Technologies (On-Site)","evidenceRecords":1},{"name":"Code Room Hub","evidenceRecords":1},{"name":"Code Room Hub(Remote) + Netsol(NIAI Training Program)","evidenceRecords":1},{"name":"CodetoKloud (job)","evidenceRecords":1},{"name":"DarkBytes","evidenceRecords":1},{"name":"DevelopersHub Corporation","evidenceRecords":1},{"name":"Devntech","evidenceRecords":1},{"name":"Everyone Company","evidenceRecords":1},{"name":"Fly","evidenceRecords":1},{"name":"Glosix System","evidenceRecords":1},{"name":"Incubee","evidenceRecords":1},{"name":"ISPR (Inter Service Public Relations)","evidenceRecords":1},{"name":"Parallax Labs(remote)","evidenceRecords":1},{"name":"Poplone","evidenceRecords":1},{"name":"Quantum logistics","evidenceRecords":1},{"name":"Shezan International Limited","evidenceRecords":1},{"name":"Synoptix.ai","evidenceRecords":1},{"name":"System","evidenceRecords":1},{"name":"system solution","evidenceRecords":1},{"name":"Tech Orbito","evidenceRecords":1},{"name":"Technisia","evidenceRecords":1},{"name":"Techtimize.co","evidenceRecords":1},{"name":"Unilever","evidenceRecords":1},{"name":"Unisoftware","evidenceRecords":1},{"name":"UniSoftwares Pakistan","evidenceRecords":1},{"name":"Web marketing solutions","evidenceRecords":1},{"name":"Web7labs","evidenceRecords":1}];
const CONFIRMED_IP_TOTAL = 177;
function isCompleteOverviewScope(){
  return ($("monthFilter")?.value || "All") === "All" &&
    ($("facultyFilter")?.value || "All") === "All" &&
    ($("departmentFilter")?.value || "All") === "All" &&
    ($("organizationFilter")?.value || "All") === "All" &&
    !($('searchFilter')?.value || '').trim();
}
function reconciledTotal(rows) {
  // One authoritative dashboard collection: all counts use the same filtered rows.
  return rows.reduce((sum, row) => sum + rowAmount(row), 0);
}

function rowAmount(row) {
  // For a selected PRPSTFPT41 evidence organization, display its own frequency.
  // Do not allocate the entire 57-student submission to every company.
  const selectedOrg=document.getElementById('organizationFilter')?.value||'All';
  if(selectedOrg!=='All'&&isPr41Internship(row)){
    const match=pr41OrgEvidence.organizations.find(x=>x.name===selectedOrg);
    if(match)return match.evidenceRecords;
  }
  if(selectedOrg!=='All'&&isPr46Internship(row)){
    const match=pr46OrgEvidence.organizations.find(x=>x.name===selectedOrg);
    if(match)return match.records;
  }
  return row.events > 0 ? row.events : 1;
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


function escapeHtmlText(value) {
  return String(value ?? "").replace(/[&<>"']/g, c => {
    const entities = {"&":"&amp;", "<":"&lt;", ">":"&gt;", "'":"&#39;"};
    return c === '"' ? "&quot;" : entities[c];
  });
}

function renderDepartmentRanking(data) {
  const host = $("departmentRanking");
  if (!host) return;
  const entries = Object.entries(data || {});
  const totalGroups = entries.length;
  const max = Math.max(1, ...entries.map(([,n]) => Number(n)||0));
  const colors = ["#0c9478", "#2e68c7", "#7c57c2", "#de8e23", "#1c88a2", "#2f4ea8", "#0b7f6f", "#5b7be1"];
  host.innerHTML = `
    <div class="department-ranking-top">
      <div>
        <div class="department-ranking-title">Department-wise Activities</div>
        <p class="department-ranking-subtitle">Total activity volume by department</p>
      </div>
      <span class="department-ranking-badge">${totalGroups} groups</span>
    </div>
    <div class="department-ranking-list">
      ${entries.map(([name, value], index) => {
        const n = Number(value)||0;
        const width = Math.max(3, n / max * 100);
        const color = colors[index % colors.length];
        return `<button type="button" class="department-ranking-row" data-department-name="${escapeHtmlText(name)}" aria-label="${escapeHtmlText(name)}, ${n.toLocaleString()} activities. View details">
          <span class="department-ranking-rank">${String(index+1).padStart(2,'0')}</span>
          <span class="department-ranking-main">
            <span class="department-ranking-name" title="${escapeHtmlText(name)}">${escapeHtmlText(name)}</span>
            <span class="department-ranking-track"><span class="department-ranking-fill" style="--dept-color:${color};width:${width}%"></span></span>
          </span>
          <span class="department-ranking-value">${n.toLocaleString()}</span>
        </button>`;
      }).join('')}
    </div>`;
  host.querySelectorAll('[data-department-name]').forEach(btn => btn.addEventListener('click', () => openPopup('department', btn.dataset.departmentName)));
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

  if (id === "departmentChart") {
    renderDepartmentRanking(data);
    if ($(wrapper)) {
      $(wrapper).style.height = "0px";
      $(wrapper).hidden = true;
    }
    const mobileHost = $("mobileDepartmentActivities");
    if (mobileHost) {
      mobileHost.hidden = true;
      mobileHost.replaceChildren();
    }
    const canvasEl = $(id);
    if (canvasEl) canvasEl.style.display = "none";
    return;
  }

  if (
    wrapper &&
    $(wrapper)
  ) {

    $(wrapper)
      .style
      .height =
      Math.max(280, labels.length * 47) +
      "px";
  }

  // A separate, touch-friendly mobile rendering; desktop canvas remains unchanged.
  if (id === "facultyChart" || id === "departmentChart") {
    const host = $(id === "facultyChart" ? "mobileFacultyActivities" : "mobileDepartmentActivities");
    if (host) {
      host.replaceChildren();
      const colors = ["#0b6b3a", "#1556c0", "#7f2db4", "#ef7d17", "#0b7f90", "#14206d"];
      const max = Math.max(1, ...values.map(n => Number(n) || 0));
      const intro = document.createElement("div"); intro.className = "uol-mobile-chart-intro";
      intro.textContent = `${labels.length} ${key === "faculty" ? "faculties" : "departments"} · ${chartTotal.toLocaleString()} activities`;
      host.append(intro);
      labels.forEach((name, index) => {
        const n = Number(values[index]) || 0;
        const card = document.createElement("button"); card.type = "button"; card.className = "uol-mobile-chart-card";
        card.style.setProperty("--rank-color", colors[index % colors.length]);
        card.setAttribute("aria-label", `${name}, ${n} activities, ${chartPercent(n)}. View details`);
        const head = document.createElement("div"); head.className = "uol-mobile-chart-head";
        const rank = document.createElement("span"); rank.className = "uol-mobile-chart-rank"; rank.textContent = String(index + 1).padStart(2,"0");
        const title = document.createElement("strong"); title.textContent = name;
        const amount = document.createElement("span"); amount.className = "uol-mobile-chart-value"; amount.textContent = n.toLocaleString();
        head.append(rank,title,amount);
        const bottom = document.createElement("div"); bottom.className = "uol-mobile-chart-bottom";
        const track = document.createElement("span"); track.className = "uol-mobile-chart-track";
        const fill = document.createElement("span"); fill.style.width = `${n / max * 100}%`; track.append(fill);
        const pct = document.createElement("span"); pct.className = "uol-mobile-chart-pct"; pct.textContent = chartPercent(n);
        bottom.append(track,pct); card.append(head,bottom);
        card.addEventListener("click", () => openPopup(key,name)); host.append(card);
      });
    }
  }

  const visibleBarValuePlugin = {
    id: `visibleBarValue_${id}`,
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta = chart.getDatasetMeta(0);
      ctx.save();
      ctx.font = "700 12px Inter, Segoe UI, Arial, sans-serif";
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
                title: items => items.length ? String(labels[items[0].dataIndex] || "") : "",
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

                    if (id === "departmentChart") {
                      const name = String(label || "")
                        .replace(/^Department of\s+/i, "")
                        .replace(/^Faculty of\s+/i, "")
                        .replace(/^Lahore School of\s+/i, "Lahore School: ");
                      return name.length > 31 ? name.slice(0, 28).trimEnd() + "…" : name;
                    }
                    return wrapLabel(label, 32);
                  },

                font: {

                  size: id === "departmentChart" ? 10 : 12
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

  // Always display the full institutional activity-type roster in the distribution.
  // Filtering by Spring/Fall changes each count, not the list of activity categories.
  const availableTypes = unique(allRows.map(row => row.type).filter(type =>
    !isInvalidLabel(type) && norm(type) !== "total activities"
  ));
  // In complete-institution views, the confirmed internship/placement count
  // replaces the old 1,087 contribution; it is NOT added to that contribution.
  // Narrower faculty/department/semester/organization searches remain based on
  // their actual matched source rows (the 120 lack complete attribution).
  const useConfirmedIP = isCompleteOverviewScope();
  const entries = availableTypes
    .map(type => [type, type === "Internship / Placement" && useConfirmedIP
      ? CONFIRMED_IP_TOTAL : Number(grouped[type] || 0)])
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

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

  // Keep the ring uncluttered: use the central total and the detailed
  // breakdown for numeric labels rather than putting text on narrow slices.
  if ($("activityDonutTotal")) $("activityDonutTotal").textContent = activityTotal.toLocaleString();
  if ($("activityDonutCategories")) $("activityDonutCategories").textContent =
    `${labels.length} activity categories`;

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
                4,

              borderRadius:
                3,

              hoverOffset:
                8
            }
          ]
        },

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          cutout:
            "75%",

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

  const peak = Math.max(1, ...values.map(value => Number(value || 0)));
  legend.innerHTML = labels.map((label, index) => {
    const value = Number(values[index] || 0);
    const width = Math.max(1, value / peak * 100);
    return `
      <button type="button" class="activity-legend-item" data-index="${index}"
        aria-label="View ${esc(label)}: ${value.toLocaleString()} activities">
        <span class="activity-legend-rank">${String(index + 1).padStart(2, '0')}</span>
        <span class="activity-legend-content">
          <span class="activity-legend-title">
            <span class="activity-legend-dot" style="background:${colors[index]}"></span>
            <span class="activity-legend-name" title="${esc(label)}">${esc(label)}</span>
          </span>
          <span class="activity-bar-track"><span class="activity-bar-fill" style="width:${width}%;background:${colors[index]}"></span></span>
        </span>
        <span class="activity-legend-metrics">
          <strong class="activity-legend-value">${value.toLocaleString()}</strong>
          <span class="activity-legend-share">${pct(value)}</span>
        </span>
      </button>`;
  }).join("");

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
    const semester = semesterForActivity(row);

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

  for (const semester of ["Spring Semester 2026", "Fall Semester 2026"]) {
    if (!(semester in result)) result[semester] = 0;
  }
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
        type: "bar",
        data: {
          labels,
          datasets: [
            {
              data: labels.map(label => result[label]),
              backgroundColor: ["#0c896d", "#275dcc", "#8b5cf6", "#ed9a33"],
              borderRadius: 12,
              maxBarThickness: 90
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

function renderPr46CompanyEvidence(){
 const target=$('pr46ExecutiveOrganizations');if(!target)return;
 const matches=filteredRows.some(isPr46Internship);target.hidden=!matches;target.replaceChildren();if(!matches)return;
 const heading=document.createElement('h3');heading.textContent='PRPSTFPT46 · Internship — Organization breakdown';
 const info=document.createElement('p');info.textContent=`“Multiple” resolved to ${pr46OrgEvidence.organizations.length} distinct evidence-based organization names. ${pr46OrgEvidence.evidenceStudents||98} student records (${pr46OrgEvidence.namedCompanyRows||91} with company names); the submission reports ${pr46OrgEvidence.reportedStudents||102} students counted once. Organization counts are evidence mentions, not additional placement quantities.`;
 const container=document.createElement('div');container.className='pr41-executive-list';const selected=$('organizationFilter')?.value||'All';
 (pr46OrgEvidence.organizations || []).filter(x=>selected==='All'||x.name===selected).forEach((item,i)=>{const card=document.createElement('div');card.className='pr41-executive-company';const label=document.createElement('strong');label.textContent=`${i+1}. ${item.name}`;const count=document.createElement('span');count.textContent=`${item.records} evidence ${item.records===1?'record':'records'} · ${(item.records/(pr46OrgEvidence.namedCompanyRows||91)*100).toFixed(1)}% of named-company rows`;card.append(label,count);card.tabIndex=0;card.setAttribute('role','button');card.setAttribute('aria-label',`Open ${item.name}: ${item.records} evidence records`);card.onclick=()=>window.PR41CompanyPopup.show(item,{submission:'PRPSTFPT46',total:pr46OrgEvidence.namedCompanyRows||91,reported:pr46OrgEvidence.reportedStudents||102});card.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();card.click()}};container.append(card)});target.append(heading,info,container);
}
function renderPr41CompanyEvidence(){
 const target=$('pr41ExecutiveOrganizations');if(!target)return;
 const matches=filteredRows.some(isPr41Internship);target.hidden=!matches;target.replaceChildren();if(!matches)return;
 const heading=document.createElement('h3');heading.textContent='PRPSTFPT41 · Internship / Placement (IP) — Organization names';
 const info=document.createElement('p');info.textContent=`${pr41OrgEvidence.organizations.length} unique organization display entries from Section 4 evidence (repeated names grouped); ${pr41OrgEvidence.reportedStudents||57} students reported once. Counts below indicate evidence mentions, not additional internship quantities.`;
 const activityHeading=document.createElement('h4');activityHeading.className='pr41-other-heading';activityHeading.textContent='Other activities reported in Section 2';
 const activityInfo=document.createElement('p');activityInfo.textContent='Four separately reported activity rows (quantity 1 each); these are not additional internship students or placement companies.';
 const activityGrid=document.createElement('div');activityGrid.className='pr41-other-activities';
 const otherActivities=[['Seminar / Conference / Workshop','ICACS-2026','10 Feb 2026','Successfully Organized'],['Guest Lecture',"Unisoftware’s Pakistan",'21 May 2026','Successfully Organized'],['Seminar / Conference / Workshop','AWS Community','24 Jan 2026','Successfully Organized Workshop'],['IAB Meeting / Industry Consultation','Faculty of Information Technology, UOL','26 Feb 2026','FYP Expo 2026 Successfully Organized']];
 for(const [type,partner,date,outcome] of otherActivities){const item=document.createElement('article');item.className='pr41-other-activity';const kind=document.createElement('small');kind.textContent=type;const name=document.createElement('strong');name.textContent=partner;const detail=document.createElement('span');detail.textContent=`Quantity 1 · ${date}`;const status=document.createElement('em');status.textContent=outcome;item.append(kind,name,detail,status);activityGrid.append(item)}
 const container=document.createElement('div');container.className='pr41-executive-list';const selected=$('organizationFilter')?.value||'All';
 (pr41OrgEvidence.organizations || []).filter(x=>selected==='All'||x.name===selected).forEach((item,i)=>{const card=document.createElement('div');card.className='pr41-executive-company';const label=document.createElement('strong');label.textContent=`${i+1}. ${item.name}`;const count=document.createElement('span');count.textContent=`${item.evidenceRecords} evidence ${item.evidenceRecords===1?'record':'records'} · ${(item.evidenceRecords/(pr41OrgEvidence.evidenceRows||57)*100).toFixed(1)}%`; card.append(label,count);card.tabIndex=0;card.setAttribute('role','button');card.setAttribute('aria-label',`Open ${item.name}: ${item.evidenceRecords} evidence records`);card.onclick=()=>window.PR41CompanyPopup.show(item,{submission:'PRPSTFPT41',total:pr41OrgEvidence.evidenceRows||57,reported:pr41OrgEvidence.reportedStudents||57});card.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();card.click()}};container.append(card)});target.append(heading,info,container,activityHeading,activityInfo,activityGrid);
}
function render() {

  const totalActivities = reconciledTotal(filteredRows, {
    wholeCohort: isCompleteOverviewScope() &&
      ($("activityFilter")?.value || "All") === "All"
  });

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

  // The institutional roster is independent of the selected reporting semester.
  // A semester with zero activities still contains the university's faculties and departments.
  const selectedFacultyForRoster = $("facultyFilter")?.value || "All";
  const selectedDepartmentForRoster = $("departmentFilter")?.value || "All";
  const rosterFacultyCount = new Set(uolInstitutionalRoster
    .filter(r => selectedFacultyForRoster === "All" || r.faculty === selectedFacultyForRoster)
    .filter(r => selectedDepartmentForRoster === "All" || r.department === selectedDepartmentForRoster)
    .map(r => r.faculty)).size;
  const rosterDepartmentCount = new Set(uolInstitutionalRoster
    .filter(r => selectedFacultyForRoster === "All" || r.faculty === selectedFacultyForRoster)
    .filter(r => selectedDepartmentForRoster === "All" || r.department === selectedDepartmentForRoster)
    .map(r => r.department)).size;

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

  // Count DISTINCT submission references, rather than activity quantities.
  // Keep reports with no activities in the count when no activity/organization/search
  // is selected. All options use the same historical reference snapshot.
  const submissionCountLabel = $("matchedSubmissionCount");
  if (submissionCountLabel) {
    const monthSelection = $("monthFilter")?.value || "All";
    const facultySelection = $("facultyFilter")?.value || "All";
    const departmentSelection = $("departmentFilter")?.value || "All";
    const activitySelection = $("activityFilter")?.value || "All";
    const organizationSelection = $("organizationFilter")?.value || "All";
    const searchSelection = clean($("searchFilter")?.value || "");
    const baseRefs = uolSubmissionRoster.filter(row =>
      (monthSelection === "All" || semesterForActivity(row) === monthSelection) &&
      (facultySelection === "All" || row.faculty === facultySelection) &&
      (departmentSelection === "All" || row.department === departmentSelection)
    ).map(row => row.ref);
    const requireActivities = activitySelection !== "All" || organizationSelection !== "All" || Boolean(searchSelection);
    const matchedRefs = requireActivities ? new Set(filteredRows.map(row => row.ref)) : new Set(baseRefs);
    submissionCountLabel.textContent = `Matching submissions: ${matchedRefs.size.toLocaleString("en-US")} / ${window.UOLSnapshots.cohorts[window.UOLSnapshots.active].length}`;
  }

  const selectedActivityType = $("activityFilter")?.value || "All";

  const totalActivitiesKpiCard = $("totalActivitiesKpiCard");
  if (totalActivitiesKpiCard) {
    // Show the overall total only when viewing all activity types.
    // Once a specific activity type is selected, this KPI is hidden.
    totalActivitiesKpiCard.hidden = selectedActivityType !== "All";
  }

  const internshipKpiCard = $("internshipKpiCard");
  if (internshipKpiCard) {
    const showActivityCard = selectedActivityType !== "All";
    internshipKpiCard.hidden = !showActivityCard;
    if (showActivityCard) {
      const activityIcons = {
        "Alumni Talk / Mentoring": ["🎓", "linear-gradient(90deg,#0a8a6b,#57c1a8)", "linear-gradient(135deg,#0a8a6b,#18a97a)", "Selected activity total for the current view"],
        "Community Engagement": ["🤝", "linear-gradient(90deg,#0d8c75,#4cc9a6)", "linear-gradient(135deg,#0d8c75,#16b38e)", "Selected activity total for the current view"],
        "Conference": ["🎤", "linear-gradient(90deg,#4f46e5,#818cf8)", "linear-gradient(135deg,#4f46e5,#6366f1)", "Selected activity total for the current view"],
        "Curriculum Feedback Session": ["📝", "linear-gradient(90deg,#7c3aed,#c084fc)", "linear-gradient(135deg,#7c3aed,#a855f7)", "Selected activity total for the current view"],
        "Curriculum Feedback": ["📝", "linear-gradient(90deg,#7c3aed,#c084fc)", "linear-gradient(135deg,#7c3aed,#a855f7)", "Selected activity total for the current view"],
        "Guest Lecture / GLIT": ["🎙️", "linear-gradient(90deg,#1d4ed8,#60a5fa)", "linear-gradient(135deg,#1d4ed8,#3b82f6)", "Selected activity total for the current view"],
        "IAB / Industry Consultation": ["🏭", "linear-gradient(90deg,#b45309,#f59e0b)", "linear-gradient(135deg,#b45309,#d97706)", "Selected activity total for the current view"],
        "Industrial Visit / IV": ["🚌", "linear-gradient(90deg,#0f766e,#2dd4bf)", "linear-gradient(135deg,#0f766e,#14b8a6)", "Selected activity total for the current view"],
        "Internship / Placement": ["💼", "linear-gradient(90deg,#7c3aed,#b084ff)", "linear-gradient(135deg,#7c3aed,#a855f7)", "IP + Internship/Placement"],
        "MoU / MoU Signing": ["✍️", "linear-gradient(90deg,#6d28d9,#a78bfa)", "linear-gradient(135deg,#6d28d9,#8b5cf6)", "Selected activity total for the current view"],
        "Others": ["✨", "linear-gradient(90deg,#475569,#94a3b8)", "linear-gradient(135deg,#475569,#64748b)", "Selected activity total for the current view"],
        "Planned": ["📅", "linear-gradient(90deg,#c2410c,#fdba74)", "linear-gradient(135deg,#c2410c,#ea580c)", "Selected activity total for the current view"],
        "Remote": ["💻", "linear-gradient(90deg,#0369a1,#67e8f9)", "linear-gradient(135deg,#0369a1,#0ea5e9)", "Selected activity total for the current view"],
        "Research Collaboration": ["🔬", "linear-gradient(90deg,#1e40af,#93c5fd)", "linear-gradient(135deg,#1e40af,#3b82f6)", "Selected activity total for the current view"],
        "Seminar": ["📘", "linear-gradient(90deg,#15803d,#86efac)", "linear-gradient(135deg,#15803d,#22c55e)", "Selected activity total for the current view"],
        "Workshop": ["🛠️", "linear-gradient(90deg,#b45309,#fbbf24)", "linear-gradient(135deg,#b45309,#f59e0b)", "Selected activity total for the current view"]
      };
      const [symbol, gradient, iconBg, footText] = activityIcons[selectedActivityType] || ["📊", "linear-gradient(90deg,#2563eb,#60a5fa)", "linear-gradient(135deg,#2563eb,#3b82f6)", "Selected activity total for the current view"];
      const labelNode = $("activityTypeKpiLabel");
      const valueNode = $("activityTypeKpiValue");
      const footNode = $("activityTypeKpiFoot");
      const iconNode = $("activityTypeKpiIcon");
      if (labelNode) labelNode.textContent = selectedActivityType;
      // Display the same owner-confirmed 120 as Section 2 for the
      // unfiltered internship scope. Do not fabricate filtered allocations.
      const confirmedIpView = selectedActivityType === "Internship / Placement" &&
        ($("facultyFilter")?.value || "All") === "All" &&
        ($("departmentFilter")?.value || "All") === "All" &&
        ($("monthFilter")?.value || "All") === "All" &&
        ($("organizationFilter")?.value || "All") === "All" &&
        !($('searchFilter')?.value || '').trim();
      if (valueNode) valueNode.textContent = (confirmedIpView ? CONFIRMED_IP_TOTAL : totalActivities).toLocaleString("en-US");
      if (footNode) footNode.textContent = confirmedIpView
        ? "Internship / Placement"
        : selectedActivityType === "Internship / Placement"
          ? "Internship / Placement"
          : footText;
      if (iconNode) iconNode.textContent = symbol;
      internshipKpiCard.style.setProperty("--activity-gradient", gradient);
      internshipKpiCard.style.setProperty("--activity-icon-bg", iconBg);
    }
  }

  setText(
    "kActivities",
    totalActivities
      .toLocaleString()
  );
setText(
    "kFaculties",
    faculties.size
      .toLocaleString()
  );

  setText(
    "kDepartments",
    new Set(filteredRows.map(row => row.department).filter(v => !isInvalidLabel(v))).size
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
    (() => {
      const result = new Map(Object.entries(group("faculty")));
      const selectedFaculty = $("facultyFilter")?.value || "All";
      for (const item of uolInstitutionalRoster) {
        if (selectedFaculty === "All" || item.faculty === selectedFaculty) {
          if (!result.has(item.faculty)) result.set(item.faculty, 0);
        }
      }
      return Object.fromEntries([...result.entries()].sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0])));
    })(),
    "faculty",
    "facultyWrap"
  );

  barChart(
    "departmentChart",
    (()=>{const groups=group("department");const asMap=new Map(Object.entries(groups));const faculty=$("facultyFilter")?.value||"All";
      for(const r of uolInstitutionalRoster)if(faculty==="All"||r.faculty===faculty)if(!asMap.has(r.department))asMap.set(r.department,0);
      return Object.fromEntries([...asMap.entries()].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0])));})(),
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

  const result = await window.UOLExcel.load('activities', paths);

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

  // Combine Planned and Remote under Internship / Placement for reporting.
  // Preserve the original type and remarks for record-level auditing.
  allRows = allRows.map(row => {
    const original = norm(row.originalType);
    const remote = /\bmode\s*:\s*remote\b/i.test(clean(row.remarks));
    const merge = remote || row.type === "Planned" || original === "planned" || original === "remote";
    return merge ? { ...row, type: "Internship / Placement" } : row;
  });

  // Planned section-4 records join Internship / Placement after duplicate checks.
  // Their source classification remains available in originalType and remarks.
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
        return { ...row, type: "Internship / Placement" };
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

  // Match Section 2's validated Internship/Placement reporting rule.
  // PRPSTFPT39 contains a quantity-one record with no named organization.
  // Keep the record in the embedded source workbook for audit, but do not
  // include it in dashboard activity counts or semester/faculty distributions.
  allRows = allRows.filter(row => !(
    clean(row.ref).toUpperCase() === 'PRPSTFPT39' &&
    row.type === 'Internship / Placement' &&
    Number(row.events) === 1 &&
    isInvalidLabel(clean(row.organization))
  ));

  // Authoritative correction for PRPSTFPT77 (Diet & Nutritional Sciences).
  // Older Excel versions can still carry 190 even if the dropdown was
  // corrected to 1. Normalize the loaded record before any KPI, chart,
  // popup, facet, or PDF aggregation uses it.
  allRows = allRows.map(row => (
    String(row.ref || '').trim().toUpperCase() === 'PRPSTFPT77' &&
    row.type === 'Internship / Placement' &&
    /UOTH and Social Security Hospital Multan Road/i.test(String(row.organization || ''))
      ? { ...row, events: 1, raw: { ...row.raw, 'How Many': 1 } }
      : row
  ));

  // Apply the selected reference-cohort to all activity KPIs and chart data.
  allRows = allRows.filter(row => window.UOLSnapshots.contains(row.ref));

  // Reconcile the legacy internship worksheet *once* into the authoritative
  // 120 source records and the 57 distinct PRPSTFPT41 audit mentions.
  // This prevents inflated legacy quantities from leaking into dropdowns,
  // KPIs, charts, organization filters, or exported filtered datasets.
  const otherActivities = allRows.filter(row => row.type !== 'Internship / Placement');
  const internshipRecords = PDF_INTERNSHIP_SOURCE_ROWS.map(source => ({
    ...source,
    month: source.reportingMonth || '',
    events: Number(source.events) || 1,
    raw: {
      'Ref #': source.ref,
      'Faculty': source.faculty,
      'Department': source.department,
      'Separated Activity Type': 'Internship / Placement',
      'Original Activity Type': source.originalType,
      'How Many': Number(source.events) || 1,
      'Organization / Activity Title': source.organization,
      'Remarks / Result': source.remarks
    }
  })).filter(row => window.UOLSnapshots.contains(row.ref));
  const companyRows = (pr41OrgEvidence.organizations || []).map(company => ({
    ref: 'PRPSTFPT41',
    faculty: 'Faculty of Information Technology',
    department: 'Department of Software Engineering',
    person: '', designation: '', month: '',
    type: 'Internship / Placement', originalType: 'IP',
    events: Number(company.evidenceRecords) || 0,
    organization: company.name, planned: '', scheduled: '', remarks: '',
    raw: { 'Ref #': 'PRPSTFPT41', 'Faculty': 'Faculty of Information Technology',
      'Department': 'Department of Software Engineering', 'How Many': Number(company.evidenceRecords) || 0,
      'Organization / Activity Title': company.name }
  })).filter(row => row.events > 0 && window.UOLSnapshots.contains(row.ref));
  allRows = [...otherActivities, ...internshipRecords, ...companyRows];

  // Separate section dashboards provide the reporting and MoU charts.

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

    try { const response = await fetch('data/pr41-internship-evidence.json',{cache:'no-store'}); if(response.ok) pr41OrgEvidence = await response.json(); } catch(error) { console.warn('PR41 organization evidence unavailable', error); }
    try { const response=await fetch('data/PRPSTFPT46_Unique_Organizations.json',{cache:'no-store'}); if(response.ok) pr46OrgEvidence=await response.json(); } catch(error) { console.warn('PR46 organization evidence unavailable',error); }
    await Promise.all([loadActivityWorkbook(),loadInstitutionalRoster()]);

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
        .textContent = `Connected · ${allRows.length.toLocaleString()} records`;
      $("connection").title = `${totalActivities.toLocaleString()} activities · ${faculties.size} faculties · ${departments.size} departments · ${ipSummary.total.toLocaleString()} Internship / Placement students`;
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


function pr46PopupBreakdown(row){
 if(!isPr46Internship(row))return '';
 const entries=pr46OrgEvidence.organizations||[];
 const named=pr46OrgEvidence.namedCompanyRows||91;
 return `<section class="pr46-record-breakdown" aria-label="PRPSTFPT46 organization evidence">
   <div class="pr46-record-heading"><div><span class="pr46-record-eyebrow">INTERNSHIP · PRPSTFPT46</span><h3>Organization breakdown</h3><p>Original company field: Multiple · Classification: Internship · Source: Section 4 · Evidence: 6a604efa01357(2).xlsx</p></div><span class="pr46-record-badge">${entries.length} unique organizations</span></div>
   <div class="pr46-record-metrics"><span><strong>${pr46OrgEvidence.reportedStudents||102}</strong> reported students</span><span><strong>${pr46OrgEvidence.evidenceStudents||98}</strong> evidence rows</span><span><strong>${named}</strong> named-company rows</span><span><strong>${pr46OrgEvidence.blankCompanyRows??7}</strong> unnamed-company students</span></div>
   <label class="pr46-record-search-label">Find an organization<input class="pr46-record-search" type="search" placeholder="Search the 72 organization names…" autocomplete="off"></label>
   <div class="pr46-record-list">${entries.map((item,i)=>`<button type="button" class="pr46-record-company" data-org-index="${i}"><span class="pr46-record-number">${String(i+1).padStart(2,'0')}</span><span class="pr46-record-name">${esc(item.name)}</span><span class="pr46-record-frequency">${Number(item.records)||0} · ${((Number(item.records)||0)/named*100).toFixed(1)}%</span></button>`).join('')}</div>
   <p class="pr46-record-note">Organization counts represent documented student rows, not completed placements or extra internships. The submission reports 102 students, but 98 appear in the supplied workbook (91 with named organizations and 7 without). Four remain unreconciled. Do not add organization totals to the reported 102.</p>
 </section>`;
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
                <strong>${isPr46Internship(row) ? "Internship (dashboard group: Internship / Placement (IP))" : displayValue(row.type)}</strong>
              </div>

              <div class="detail-field">
                <small>Original Activity Type</small>
                <strong>${isPr46Internship(row) ? "Intership (submitted spelling; normalized to Internship)" : displayValue(row.originalType)}</strong>
              </div>

              <div class="detail-field">
                <small>How Many</small>
                <strong>${Number(row.events || 0).toLocaleString()}${isPr46Internship(row) ? " reported students (98 documented)" : ""}</strong>
              </div>

              <div class="detail-field">
                <small>Semester</small>
                <strong>${displayValue(semesterForActivity(row))}</strong>
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
                <strong>${isPr46Internship(row) ? `${pr46OrgEvidence.organizations.length} named organizations (supporting evidence)` : displayValue(row.organization)}</strong>
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
            ${pr46PopupBreakdown(row)}
          </div>
        `
      )
        .join(
          ""
        );
  }

  const body=$('detailModalBody');
  if(body){
    body.querySelectorAll('.pr46-record-breakdown').forEach(section=>{
      const search=section.querySelector('.pr46-record-search');
      search?.addEventListener('input',()=>{
        const q=search.value.trim().toLocaleLowerCase();
        section.querySelectorAll('.pr46-record-company').forEach(button=>{button.hidden=!button.querySelector('.pr46-record-name').textContent.toLocaleLowerCase().includes(q)});
      });
      section.querySelectorAll('.pr46-record-company').forEach(button=>{
        button.addEventListener('click',()=>{
          const item=pr46OrgEvidence.organizations[Number(button.dataset.orgIndex)];
          if(item&&window.PR41CompanyPopup?.show)window.PR41CompanyPopup.show(item,{submission:'PRPSTFPT46',total:pr46OrgEvidence.namedCompanyRows||91,reported:pr46OrgEvidence.reportedStudents||102});
        });
      });
    });
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

          if (document.querySelector('.pr41-detail-overlay:not([hidden])')) return;
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
  const boxW = (available - gap * 3) / 4;
  const boxH = 30;

  const boxes = [
    ["SEMESTER", items["Semester"], left, startY, boxW],
    ["FACULTY", items["Faculty"], left + (boxW + gap), startY, boxW],
    ["DEPARTMENT", items["Department"], left + (boxW + gap) * 2, startY, boxW],
    ["ACTIVITY TYPE", items["Activity Type"], left + (boxW + gap) * 3, startY, boxW]
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

  // PDF must use the same confirmed full-cohort internship quantity as the
  // Executive Overview. The original internship records in the older master
  // worksheet overlap the confirmed 177 and must not be summed again.
  const completeCohort = isCompleteOverviewScope() &&
    (f.activity === "All" || f.activity === "Internship / Placement");
  let pdfRows = [...filteredRows];
  if (completeCohort) {
    // Include the confirmed internship quantity even when the older worksheet
    // has no matching internship row in this selection. This also prevents
    // an outdated 1,087-row subtotal from leaking into exported PDFs.
    pdfRows = pdfRows.filter(r => r.type !== "Internship / Placement");
    // Preserve original attributions for all 92 supplied worksheet records.
    // The unallocated balance is kept separate; no faculty/department is invented.
    const detailed = PDF_INTERNSHIP_SOURCE_ROWS.filter(r => Number(r.events) > 0)
      .map(r => ({ ...r, organization: r.organization || "Organization not stated",
        remarks: r.remarks || "-", pdfSourceDetail: true }));
    // Add the 57 distinct PRPSTFPT41 evidence entries by grouped organization.
    // They have a confirmed Faculty and Department; each company has its own quantity.
    const pr41Detailed = PR41_PDF_ORGANIZATIONS.map(x => ({
      ref: 'PRPSTFPT41', type: 'Internship / Placement', originalType: 'IP',
      faculty: 'Faculty of Information Technology', department: 'Department of Software Engineering',
      organization: x.name, events: x.evidenceRecords, remarks: 'PRPSTFPT41 audited organization evidence',
      month: '', reportingMonth: '', semester: 'Not specified', pdfSourceDetail: true
    }));
    const detailedTotal = detailed.reduce((sum,r) => sum + r.events,0) + pr41Detailed.reduce((sum,r) => sum + r.events,0);
    pdfRows.push(...detailed, ...pr41Detailed);
    const remainder = Math.max(0, CONFIRMED_IP_TOTAL - detailedTotal);
    if (remainder) pdfRows.push({
      type: "Internship / Placement", events: remainder,
      faculty: "Additional records not in uploaded worksheet", department: "Additional records not in uploaded worksheet",
      organization: "No organization details in uploaded worksheet",
      person: "", designation: "", scheduled: "", planned: "",
      remarks: "Source records requiring additional organization information.",
      month: "", reportingMonth: "", semester: "Not specified", pdfAggregateOnly: true
    });
  }
  return [{
    title: parts.length ? parts.join(" - ") : "University-wide Activity Report",
    subtitle: completeCohort ? "Internship / Placement: 120 source records plus 57 PRPSTFPT41 evidence records" : "",
    rows: pdfRows,
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
  return rows.filter(row => !row.pdfAggregateOnly).map(row => [
    semesterForActivity(row) || "-",
    reportDisplayValue(row.faculty),
    reportDisplayValue(row.department),
    reportDisplayValue(row.type),
    rowAmount(row),
    reportDisplayValue(row.organization),
    reportDisplayValue(row.person),
    reportDisplayValue(row.designation),
    reportDisplayValue(row.scheduled),
    reportDisplayValue(row.remarks)
  ]);
}

function expectedOutcomeBullets(rows) {
  const present = new Set(rows.map(row => clean(row.type)).filter(Boolean));
  const rules = [
    ["Internship / Placement", "Internship / Placement: Improve student employability and workplace readiness by expanding structured opportunities for practical training, professional exposure and recruitment. Expected results include stronger employer networks, improved job-readiness, better understanding of workplace standards and a clearer pathway from academic learning to employment."],
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
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "bold");
      doc.text("•", 19, y);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(0, 0, 0);
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
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "bold");
      doc.text("•", 19, y);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(0, 0, 0);
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
  const w = doc.internal.pageSize.getWidth();
  const h = doc.internal.pageSize.getHeight();
  doc.setFillColor(247, 250, 253); doc.rect(0, 0, w, h, "F");
  // Elegant institutional masthead, consistent with the web dashboard.
  doc.setFillColor(16, 40, 62); doc.rect(0, 0, w, 61, "F");
  doc.setFillColor(12, 150, 121); doc.rect(0, 61, w, 3, "F");
  doc.setFillColor(255, 255, 255); doc.roundedRect(15, 13, 75, 35, 3, 3, "F");
  addPdfLogo(doc, logoDataUrl, 19, 16, 67, 29);
  doc.setTextColor(167, 223, 208); doc.setFont("helvetica", "bold"); doc.setFontSize(8.4);
  doc.text("OFFICE OF EXTERNAL LINKAGES  /  ANALYTICS", 101, 21);
  doc.setTextColor(255, 255, 255); doc.setFontSize(21);
  doc.text("INSTITUTIONAL ENGAGEMENT", 101, 35);
  doc.setFontSize(16); doc.text("PERFORMANCE REPORT", 101, 45);
  doc.setFontSize(8); doc.setTextColor(208, 228, 239);
  doc.text("TENTATIVE  |  FILTERED PERFORMANCE INTELLIGENCE", 101, 53);

  doc.setTextColor(16, 44, 68); doc.setFont("helvetica", "bold"); doc.setFontSize(12);
  doc.text("REPORT SCOPE", 15, 81);
  doc.setDrawColor(216, 227, 236); doc.line(15, 85, w - 15, 85);
  if (sections && sections.length) addPdfFilterBanner(doc, sections[0], 92);

  const authorityY = 148; const cw = (w - 38) / 2;
  [[15, "DEAN OSA", "Ms. Ammara Awais Raoof"], [23 + cw, "DIRECTOR OEL", "Dr. Muhammad Shafique"]].forEach(([x, role, name], i) => {
    doc.setFillColor(255, 255, 255); doc.setDrawColor(218, 231, 240);
    doc.roundedRect(x, authorityY, cw, 30, 3, 3, "FD");
    doc.setFillColor(i ? 36 : 13, i ? 102 : 146, i ? 173 : 121);
    doc.roundedRect(x, authorityY, 3, 30, 1, 1, "F");
    doc.setTextColor(91, 120, 141); doc.setFontSize(8); doc.setFont("helvetica", "bold");
    doc.text(role, x + 9, authorityY + 11);
    doc.setTextColor(17, 48, 70); doc.setFontSize(11);
    doc.text(name, x + 9, authorityY + 22);
  });
  doc.setTextColor(102, 125, 143); doc.setFontSize(8.5);
  doc.text(`Generated: ${reportGeneratedOn()}`, 15, 192);
  doc.setFillColor(16, 40, 62); doc.rect(0, h - 9, w, 9, "F");
  doc.setTextColor(255, 255, 255); doc.setFontSize(7.5);
  doc.text("THE UNIVERSITY OF LAHORE  |  CONNECT  -  COLLABORATE  -  CREATE IMPACT", w / 2, h - 3.6, { align: "center" });
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
  const w = doc.internal.pageSize.getWidth();
  doc.setFillColor(255, 255, 255); doc.rect(0, 0, w, 57, "F");
  doc.setFillColor(16, 40, 62); doc.rect(0, 0, w, 30, "F");
  doc.setFillColor(12, 150, 121); doc.rect(0, 30, w, 2.2, "F");
  doc.setFillColor(255, 255, 255); doc.roundedRect(12, 5, 57, 20, 2, 2, "F");
  addPdfLogo(doc, logoDataUrl, 14.5, 6.2, 52, 17.5);
  doc.setFont("helvetica", "bold"); doc.setTextColor(255, 255, 255);
  doc.setFontSize(11.2); doc.text("INSTITUTIONAL ENGAGEMENT", 76, 12);
  doc.setFontSize(7.7); doc.setTextColor(179, 220, 213);
  doc.text("UNIVERSITY OF LAHORE  /  OFFICE OF EXTERNAL LINKAGES", 76, 20);
  doc.setFontSize(7.1); doc.setTextColor(255, 255, 255);
  doc.text(label, w - 12, 13, { align: "right" });
  doc.setTextColor(23, 51, 76); doc.setFont("helvetica", "bold"); doc.setFontSize(14.5);
  const title = doc.splitTextToSize(section.title || "Performance Report", 184);
  doc.text(title.slice(0, 1), 14, 42);
  doc.setFont("helvetica", "normal"); doc.setFontSize(8.2); doc.setTextColor(0, 0, 0);
  doc.text(doc.splitTextToSize(section.subtitle || "", 190).slice(0, 1), 14, 49);
  doc.text(`Report ${sectionIndex + 1} / ${sectionCount}`, w - 13, 41, { align: "right" });
  doc.text(`Generated ${reportGeneratedOn()}`, w - 13, 48.5, { align: "right" });
  doc.setDrawColor(221, 231, 239); doc.line(14, 53, w - 14, 53);
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
    doc.setTextColor(16, 55, 80);
    doc.text("THE UNIVERSITY OF LAHORE  |  Office of External Linkages", 14, height - 9.5);
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

// Organizations shown alongside each summary category; all names remain
// available in the full organization-wise table and detailed records.
function reportOrganizationsForGroup(rows, key, label) {
  const names = [...new Set(rows
    .filter(row => row[key] === label)
    .map(row => clean(row.organization))
    .filter(name => !isInvalidLabel(name)))].sort((a,b) => a.localeCompare(b));
  if (!names.length) return "-";
  const preview = names.slice(0, 5).join("; ");
  return names.length > 5 ? `${preview}; +${names.length - 5} more (see Organization-wise Activities)` : preview;
}

// Distinct recorded remarks for each summary group. Keep summary cells concise;
// full source remarks are included, untruncated, in Detailed Records.
function reportRemarksForGroup(rows, key, label) {
  const remarks = [...new Set(rows
    .filter(row => clean(row[key]) === clean(label))
    .map(row => clean(row.remarks))
    .filter(value => !isInvalidLabel(value)))];
  if (!remarks.length) return "-";
  const shown = remarks.slice(0, 2).map(value =>
    value.length > 165 ? `${value.slice(0, 162)}...` : value);
  return shown.join("; ") + (remarks.length > 2 ? `; +${remarks.length - 2} more (see Detailed Records)` : "");
}

function reportMonthBreakdown(rows) {
  const grouped = {};

  rows.forEach(row => {
    const label = semesterForActivity(row);
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

// Print-first, vector-based report graphs. All values are derived from the
// exact rows belonging to the current report section (including active filters).
// Unlike screenshots of dashboard canvases, these are sharp in saved PDFs.
function pdfGraphEntries(section, title) {
  const rows = section.rows || [];
  if (title === "Faculty-wise Activities") return reportGroupBy(rows, "faculty");
  if (title === "Department-wise Activities") return reportGroupBy(rows, "department");
  if (title === "Activity Type Distribution") return activityBreakdown(rows);
  if (title === "Semester Trend") return reportMonthBreakdown(rows);
  return [];
}

function pdfGraphColor(index) {
  const palette = [
    [10, 112, 79], [29, 90, 184], [111, 63, 185], [229, 120, 43],
    [15, 128, 145], [34, 58, 114], [198, 76, 110], [184, 137, 30],
    [35, 124, 100], [74, 113, 206]
  ];
  return palette[index % palette.length];
}

function pdfGraphText(doc, value, x, y, opts = {}) {
  doc.setFont("helvetica", opts.bold ? "bold" : "normal");
  doc.setFontSize(opts.size || 9);
  doc.setTextColor(...(opts.color || [30, 53, 76]));
  doc.text(String(value), x, y, opts.align ? { align: opts.align } : undefined);
}

function drawPdfChartPage(doc, item, title = "Graph", startY = 58, section = null) {
  const graphTitle = item?.title || title;
  const rows = section?.rows || [];
  const entries = pdfGraphEntries({ rows }, graphTitle);
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  const left = 16, right = W - 16;
  const footer = H - 22;
  const total = entries.reduce((sum, entry) => sum + Number(entry[1] || 0), 0);

  pdfGraphText(doc, graphTitle, left, startY + 5, { size: 16, bold: true, color: [13, 45, 70] });
  doc.setFillColor(16, 150, 118);
  doc.roundedRect(left, startY + 9, 47, 1.6, .7, .7, "F");
  pdfGraphText(doc, `${entries.length} ${graphTitle === "Semester Trend" ? "semesters" : graphTitle === "Activity Type Distribution" ? "activity categories" : "groups"}  |  ${total.toLocaleString("en-US")} reported activities`, right, startY + 5, { size: 9, color: [102, 123, 142], align: "right" });

  if (!entries.length) {
    pdfGraphText(doc, "No data for the selected report filters.", left, startY + 25, { size: 12 });
    return;
  }

  if (graphTitle === "Semester Trend") {
    // Separate bars, not a line connecting only two categorical semesters.
    const max = Math.max(1, ...entries.map(v => Number(v[1]) || 0));
    const chartY = startY + 31;
    const chartH = Math.max(55, footer - chartY - 23);
    const n = entries.length;
    const plotW = right - left - 48;
    const colW = plotW / n;
    entries.forEach(([name, value], i) => {
      const x = left + 23 + i * colW + colW / 2;
      const height = Math.max(1.5, chartH * value / max);
      const barW = Math.min(57, colW * .38);
      doc.setFillColor(...pdfGraphColor(i));
      doc.roundedRect(x - barW / 2, chartY + chartH - height, barW, height, 2, 2, "F");
      pdfGraphText(doc, Number(value).toLocaleString("en-US"), x, chartY + chartH - height - 5, { size: 15, bold: true, align: "center" });
      const lines = doc.splitTextToSize(String(name), Math.max(35, colW - 12));
      doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(42, 65, 86);
      doc.text(lines.slice(0, 2), x, chartY + chartH + 8, { align: "center" });
    });
    doc.setDrawColor(218, 227, 235); doc.setLineWidth(.4);
    doc.line(left + 16, chartY + chartH, right - 16, chartY + chartH);
    return;
  }

  if (graphTitle === "Activity Type Distribution") {
    // A fully ranked list is far more legible in print than a tiny 15-color legend.
    const top = startY + 22;
    const rowH = Math.min(9.2, (footer - top - 4) / entries.length);
    const max = Math.max(1, ...entries.map(v => Number(v[1]) || 0));
    const labelW = 81, graphX = left + labelW + 3, graphW = right - graphX - 25;
    entries.forEach(([name, value], i) => {
      const y = top + i * rowH;
      if (i % 2 === 0) { doc.setFillColor(246, 249, 252); doc.rect(left, y - 4.8, right - left, rowH, "F"); }
      pdfGraphText(doc, String(name).length > 38 ? String(name).slice(0, 36) + "..." : name, left + 2, y + .6, { size: 8.6, bold: i < 3 });
      doc.setFillColor(229, 236, 241); doc.roundedRect(graphX, y - 3.1, graphW, 4.2, 1, 1, "F");
      doc.setFillColor(...pdfGraphColor(i)); doc.roundedRect(graphX, y - 3.1, Math.max(.6, graphW * value / max), 4.2, 1, 1, "F");
      pdfGraphText(doc, Number(value).toLocaleString("en-US"), right - 1, y + .6, { size: 9, bold: true, align: "right" });
    });
    return;
  }

  // PDF rendition of the responsive ranked cards. Counts and shares are
  // computed from the same filtered report rows used by desktop/mobile.
  const isDepartment = graphTitle === "Department-wise Activities";
  const columns = isDepartment && entries.length > 15 ? 2 : 1;
  const gap = 9;
  const colWidth = (right - left - gap * (columns - 1)) / columns;
  const perCol = Math.ceil(entries.length / columns);
  const top = startY + 20;
  const available = footer - top;
  const rowH = Math.min(columns === 2 ? 9.1 : 12.0, available / Math.max(1,perCol));
  const max = Math.max(1, ...entries.map(v => Number(v[1]) || 0));
  entries.forEach(([name, value], i) => {
    const col = Math.floor(i / perCol);
    const ri = i % perCol;
    const x = left + col * (colWidth + gap);
    const y = top + ri * rowH;
    const n = Number(value) || 0;
    const share = total ? `${(n / total * 100).toFixed(1)}%` : "0.0%";
    const color = pdfGraphColor(i);
    // Card-like rows with rank, legible labels, count, percentage and bar.
    doc.setFillColor(i % 2 ? 250 : 245, 248, 252);
    doc.setDrawColor(229, 237, 244);
    doc.roundedRect(x, y, colWidth - 1, rowH - .65, 1.4, 1.4, "FD");
    doc.setFillColor(...color);
    doc.roundedRect(x, y, 1.2, rowH - .65, .5, .5, "F");
    const fs = columns === 2 ? 6.3 : 8.5;
    const labelX = x + (columns === 2 ? 13 : 17);
    const valueArea = columns === 2 ? 25 : 34;
    const labelWidth = colWidth - (labelX - x) - valueArea - 3;
    pdfGraphText(doc, String(i + 1).padStart(2, "0"), x + 3, y + 3.3, {size:columns === 2 ? 6 : 7.4, bold:true, color});
    doc.setFont("helvetica", "bold"); doc.setFontSize(fs); doc.setTextColor(32, 57, 78);
    const lines = doc.splitTextToSize(String(name), labelWidth).slice(0, 2);
    doc.text(lines, labelX, y + (lines.length > 1 ? 2.6 : 3.6));
    pdfGraphText(doc, `${n.toLocaleString("en-US")}  (${share})`, x + colWidth - 4, y + 3.4,
      {size:columns === 2 ? 6.2 : 8, bold:true, align:"right"});
    const barX = labelX, barY = y + rowH - 3.7;
    const barWidth = colWidth - (barX - x) - 5;
    doc.setFillColor(227, 235, 241);
    doc.roundedRect(barX, barY, barWidth, 1.65, .8, .8, "F");
    if (n > 0) {doc.setFillColor(...color);doc.roundedRect(barX, barY, Math.max(.5, barWidth * n/max), 1.65, .8, .8, "F");}
  });
  pdfGraphText(doc, "Share calculated within the selected report records · ranked by reported activity volume", left, footer + 4,
    {size:7, color:[114, 133, 150]});
}

// Central table renderer: one consistent, print-friendly visual system for
// faculty, department, activity and detailed-record pages.
function addModernPdfTable(doc, config) {
  if (typeof doc.autoTable !== "function") {
    throw new Error("The PDF table library is missing (jspdf-autotable).");
  }
  const pageWidth = doc.internal.pageSize.getWidth();
  const usableWidth = pageWidth - 28;
  const inputWidths = config.widths || [];
  const widthTotal = inputWidths.reduce((a, b) => a + b, 0) || 1;
  // Calculate widths from the available page width; avoid overflowing A4.
  const columnStyles = Object.fromEntries(inputWidths.map((w, i) => [i, {
    cellWidth: usableWidth * w / widthTotal,
    halign: (config.numericColumns || []).includes(i) ? "right" : "left",
    fontStyle: (config.numericColumns || []).includes(i) ? "bold" : "normal"
  }]));
  const compact = !!config.compact;
  const count = config.rows.length;
  const totalText = `${count.toLocaleString("en-US")} ${count === 1 ? "entry" : "entries"}`;
  const header = config.onPage;
  doc.autoTable({
    startY: config.startY || 60,
    head: [config.headers],
    body: config.rows,
    theme: "grid",
    showHead: "everyPage",
    rowPageBreak: "avoid",
    margin: { left: 14, right: 14, top: 58, bottom: 18 },
    tableWidth: usableWidth,
    styles: {
      font: "helvetica", fontSize: compact ? 7.4 : 9.2,
      cellPadding: compact ? {top: 2.8, bottom: 2.8, left: 2.2, right: 2.2} : {top: 4.1, bottom: 4.1, left: 4, right: 4},
      overflow: "linebreak", valign: "middle", textColor: [8, 25, 42],
      lineColor: [218, 228, 235], lineWidth: 0.16
    },
    headStyles: {
      fillColor: [12, 45, 69], textColor: [255, 255, 255],
      fontStyle: "bold", fontSize: compact ? 7.3 : 9.1,
      cellPadding: compact ? 3 : 4.5, lineColor: [12, 45, 69], lineWidth: .2,
      minCellHeight: compact ? 12 : 14
    },
    bodyStyles: { minCellHeight: compact ? 10 : 13 },
    alternateRowStyles: { fillColor: [245, 249, 251] },
    columnStyles,
    didParseCell: function (data) {
      if (data.section === "body" && data.column.index === 0) {
        data.cell.styles.fontStyle = "bold";
      }
    },
    didDrawPage: function (data) {
      if (typeof header === "function" && data.pageNumber > 1) header();
      const W = doc.internal.pageSize.getWidth();
      doc.setFont("helvetica", "normal"); doc.setFontSize(8);
      doc.setTextColor(16, 39, 59);
      doc.text(totalText, W - 14, doc.internal.pageSize.getHeight() - 12, {align:"right"});
    }
  });
  // A small final summary below breakdown tables, if the page has space.
  if (config.summaryLabel && doc.lastAutoTable) {
    const y = doc.lastAutoTable.finalY + 8;
    const H = doc.internal.pageSize.getHeight();
    if (y + 11 < H - 17) {
      doc.setFillColor(232, 246, 240);
      doc.setDrawColor(191, 223, 209);
      doc.roundedRect(14, y, usableWidth, 11, 2, 2, "FD");
      doc.setFont("helvetica", "bold"); doc.setTextColor(10, 94, 71); doc.setFontSize(9);
      doc.text(config.summaryLabel, 18, y + 7);
      doc.text(config.summaryValue || "", pageWidth - 18, y + 7, {align:"right"});
    }
  }
}

// Faculty/department organization names come from the separate, attributed
// source worksheet. They are evidence only and do not alter the confirmed 177 KPI.
async function pdfInternshipOrganizationRows() {
  const f = getCurrentFilterSummary();
  const standard = value => String(value || "").replace(/\s+/g, " ").trim();
  const attributedRows = [...PDF_INTERNSHIP_SOURCE_ROWS, ...PR41_PDF_ORGANIZATIONS.map(x => ({
    ref: 'PRPSTFPT41', faculty: 'Faculty of Information Technology',
    department: 'Department of Software Engineering', organization: x.name,
    events: x.evidenceRecords, remarks: 'PRPSTFPT41 audited evidence'
  }))];
  return attributedRows.map(r => ({
    faculty: standard(r.faculty), department: standard(r.department),
    organization: standard(r.organization), ref: standard(r.ref),
    quantity: Number(r.events) || 0, remarks: standard(r.remarks),
    activity: "Internship / Placement"
  })).filter(r => r.quantity > 0 && r.organization && !isInvalidLabel(r.organization))
    .filter(r => f.faculty === "All" || standard(f.faculty) === r.faculty)
    .filter(r => f.department === "All" || standard(f.department) === r.department)
    .filter(r => f.activity === "All" || f.activity === "Internship / Placement")
    .filter(r => f.organization === "All" || standard(f.organization) === r.organization)
    .filter(r => !f.search || [r.organization,r.faculty,r.department,r.ref,r.remarks].join(" ").toLowerCase().includes(f.search.toLowerCase()));
}

function addPdfInternshipOrganizationAppendix(doc, section, index, sectionCount, logoDataUrl, sourceRows) {
  if (!sourceRows.length) return;
  // Each record retains its original faculty and department; 127 is not
  // distributed among them because those full allocations remain unverified.
  for (const spec of [
    { key: "faculty", label: "FACULTY-WISE INTERNSHIP ORGANIZATIONS", col: "FACULTY" },
    { key: "department", label: "DEPARTMENT-WISE INTERNSHIP ORGANIZATIONS", col: "DEPARTMENT" }
  ]) {
    doc.addPage();
    const reportSection = { ...section, subtitle: "Organization names and quantities from the linked internship worksheet" };
    const header = () => addPdfHeader(doc, reportSection, index, sectionCount, spec.label, logoDataUrl);
    header();
    const list = sourceRows.slice().sort((a,b) => a[spec.key].localeCompare(b[spec.key]) || a.organization.localeCompare(b.organization));
    addModernPdfTable(doc, {
      startY: 60,
      headers: [spec.col, "ORGANIZATION NAME", "SUBMISSION REF", "REPORTED QUANTITY"],
      rows: list.map(r => [r[spec.key] || "Not provided", r.organization, r.ref || "-", r.quantity.toLocaleString("en-US")]),
      numericColumns: [3], widths: [66, 133, 43, 27], compact: true, onPage: header,
      summaryLabel: "NAMED SOURCE-RECORD QUANTITY",
      summaryValue: list.reduce((sum,r) => sum+r.quantity,0).toLocaleString("en-US")
    });
  }
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
  const attributedInternshipRows = await pdfInternshipOrganizationRows();

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

    // Executive KPI cards: the same metrics as the dashboard, without a crowded grid table.
    const cardGap = 6;
    const cardW = (doc.internal.pageSize.getWidth() - 28 - cardGap * 3) / 4;
    const cards = [["ACTIVITY VOLUME", sectionTotal.toLocaleString(), "Current filtered selection"],
      ...meta.filter(([label]) => label !== "Total Activities").map(([label, value]) => [label.toUpperCase(), String(value), "Current filtered selection"])];
    cards.forEach(([label, value, caption], i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const x = 14 + col * (cardW + cardGap), y = summaryStartY + row * 39;
      doc.setFillColor(249, 252, 254); doc.setDrawColor(218, 228, 236);
      doc.roundedRect(x, y, cardW, 33, 3, 3, "FD");
      doc.setFillColor(col === 0 ? 15 : col === 1 ? 27 : col === 2 ? 66 : 16,
                       col === 0 ? 146 : col === 1 ? 99 : col === 2 ? 91 : 128,
                       col === 0 ? 115 : col === 1 ? 181 : col === 2 ? 173 : 171);
      doc.rect(x, y, cardW, 2.4, "F");
      doc.setFont("helvetica", "bold"); doc.setTextColor(98, 124, 144); doc.setFontSize(7.3);
      doc.text(doc.splitTextToSize(label, cardW - 8).slice(0, 2), x + 4, y + 10);
      doc.setFontSize(17.5); doc.setTextColor(16, 43, 67);
      doc.text(String(value), x + 4, y + 24);
      doc.setFont("helvetica", "normal"); doc.setFontSize(6.7); doc.setTextColor(139, 158, 172);
      doc.text(caption, x + 4, y + 30);
    });


    // ---------- MODERN INSTITUTIONAL BREAKDOWN TABLES ----------
    const breakdownTables = [
      { title: "FACULTY-WISE ACTIVITIES", subtitle: "Faculty-wise activity summary with organizations", firstLabel: "FACULTY", groupKey: "faculty", entries: facultyBreakdown },
      { title: "ACTIVITY BREAKDOWN", subtitle: "Activity type summary with organizations", firstLabel: "ACTIVITY TYPE", groupKey: "type", entries: breakdown },
      { title: "DEPARTMENT-WISE ACTIVITIES", subtitle: "Department-wise activity summary with organizations", firstLabel: "DEPARTMENT", groupKey: "department", entries: departmentBreakdown }
    ];
    for (const spec of breakdownTables) {
      if (!spec.entries.length) continue;
      doc.addPage();
      const tableSection = { ...section, subtitle: spec.subtitle };
      const header = () => addPdfHeader(doc, tableSection, index, sections.length, spec.title, logoDataUrl);
      header();
      addModernPdfTable(doc, {
        startY: 60,
        headers: [spec.firstLabel, "ORGANIZATION(S)", "REMARKS / RESULT", "TOTAL ACTIVITIES", "SHARE"],
        rows: spec.entries.map(([name, amount]) => [
          name, reportOrganizationsForGroup(section.rows, spec.groupKey, name),
          reportRemarksForGroup(section.rows, spec.groupKey, name),
          Number(amount).toLocaleString("en-US"), reportPercentage(amount, sectionTotal)
        ]),
        numericColumns: [3, 4],
        widths: [51, 91, 85, 24, 18],
        compact: true,
        onPage: header,
        summaryLabel: "TOTAL REPORTED",
        summaryValue: sectionTotal.toLocaleString("en-US")
      });
    }

    // ---------- FULL ORGANIZATION-WISE TABLE ----------
    // Names are sourced from the same filtered records as the report.
    const organizationBreakdown = reportGroupBy(section.rows, "organization");
    if (organizationBreakdown.length) {
      doc.addPage();
      const orgSection = { ...section, subtitle: "Full organization-wise activities and participation" };
      const orgHeader = () => addPdfHeader(doc, orgSection, index, sections.length, "ORGANIZATION-WISE ACTIVITIES", logoDataUrl);
      orgHeader();
      addModernPdfTable(doc, {
        startY: 60,
        headers: ["ORGANIZATION NAME", "REMARKS / RESULT", "TOTAL ACTIVITIES"],
        rows: organizationBreakdown.map(([name, amount]) => [
          name, reportRemarksForGroup(section.rows, "organization", name),
          Number(amount).toLocaleString("en-US")
        ]),
        numericColumns: [2],
        widths: [94, 143, 32],
        onPage: orgHeader,
        summaryLabel: "TOTAL WITH NAMED ORGANIZATIONS",
        summaryValue: organizationBreakdown.reduce((n, entry) => n + entry[1], 0).toLocaleString("en-US")
      });
    }

    // Source-attributed internship organization names, by faculty and department.
    addPdfInternshipOrganizationAppendix(doc, section, index, sections.length, logoDataUrl, attributedInternshipRows);

    // ---------- EXPECTED OUTCOMES ----------
    doc.addPage();
    addPdfExpectedOutcomesPage(doc, section, index, sections.length, logoDataUrl);

    // ---------- REMARKS ----------
    doc.addPage();
    addPdfRemarksPage(doc, section, index, sections.length, logoDataUrl);

    // ---------- VECTOR GRAPH PAGES: ONE GRAPH PER PAGE ----------
    const graphTitles = [
      "Faculty-wise Activities", "Activity Type Distribution",
      "Department-wise Activities", "Semester Trend"
    ];
    const orderedCharts = graphTitles
      .map(title => ({ title, entries: pdfGraphEntries(section, title) }))
      .filter(item => item.entries.length);

    orderedCharts.forEach((chartItem, chartIndex) => {
      doc.addPage();
      const graphSection = {
        ...section,
        subtitle: `Graph page ${chartIndex + 1} of ${orderedCharts.length}`
      };
      addPdfHeader(doc, graphSection, index, sections.length, "GRAPHICAL SUMMARY", logoDataUrl);
      drawPdfChartPage(doc, chartItem, chartItem.title, 58, section);
    });

    // ---------- DETAILED RECORDS ----------
    // Confirmed internship aggregates are not individual evidence records.
    if (!detailRows(section.rows).length) continue;
    doc.addPage();
    const detailSection = { ...section, subtitle: `${section.subtitle || ""} - Detailed Records` };
    addPdfHeader(doc, detailSection, index, sections.length, "DETAILED RECORDS", logoDataUrl);
    const detailsStartY = 58;

    addModernPdfTable(doc, {
      startY: 58,
      headers: ["SEMESTER", "FACULTY", "DEPARTMENT", "ACTIVITY TYPE", "HOW MANY", "ORGANIZATION", "FOCAL PERSON", "DESIGNATION", "SCHEDULED", "REMARKS / RESULT"],
      rows: detailRows(section.rows),
      numericColumns: [4],
      widths: [20, 27, 30, 24, 13, 36, 22, 21, 14, 62],
      compact: true,
      onPage: () => addPdfHeader(doc, detailSection, index, sections.length, "DETAILED RECORDS", logoDataUrl)
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

  for (const [sectionIndex, section] of sections.entries()) {
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
          <div class="report-label">TENTATIVE ANALYTICS REPORT</div>
        </div>
        <h1>${esc(section.title)}</h1>
        <div class="subtitle">${esc(section.subtitle || "")}</div>
        <div class="generated-on">Generated on: ${esc(reportGeneratedOn())}</div>

        ${sectionIndex === 0 ? `<div class="print-filters">
          ${reportFilterDetails(section).filter(([label]) => label !== "Organization").map(([label,value]) => `<div class="print-filter"><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`).join("")}
        </div>` : ""}

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
      /* UOL print-ready report: clean header, card metrics, tables and one chart per page */
      @page{size:A4 landscape;margin:13mm 12mm}
      body{font-family:Arial,Helvetica,sans-serif;color:#17324a;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      .report-head{border-radius:0;background:#10314b;border-bottom:4px solid #159879;padding:16px 20px}
      .report-label{border-radius:6px;font-size:9px;letter-spacing:.09em;padding:9px 12px}
      .office{color:#c8e7e2}.university{font-size:17px;letter-spacing:.7px}
      h1{font-size:23px;color:#0e3452;margin:18px 0 4px}h2{font-size:14px;color:#12354f;margin:12px 0 10px;border-left:4px solid #0b9278;padding:7px 10px;background:#f0f7f5}
      .subtitle,.generated-on{color:#64748b;font-size:10px}
      .print-filters{grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:14px 0}
      .print-filter{border:1px solid #dbe8ec;border-top:3px solid #0b9278;border-radius:8px;background:#f8fcfb;padding:11px 12px;min-height:57px}
      .print-filter span{color:#598274;font-size:9px;letter-spacing:.04em}.print-filter strong{font-size:11px;color:#0f2e46}
      .kpis{grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:10px 0 16px}.kpi{padding:13px;background:#f8fafc;border:1px solid #dde8ed;border-radius:9px;border-top:3px solid #2568b7}.kpi span{font-size:10px;color:#627b91}.kpi strong{display:block;font-size:22px;color:#10344f;margin-top:7px}
      .page-section{break-before:page;page-break-before:always}.graph-grid{display:block!important}.graph-card{break-before:page;page-break-before:always;break-inside:avoid;border:0;padding:0;margin:0;background:white}.graph-title{font-size:17px;color:#0e344d;padding:11px 0;border-bottom:3px solid #0b9278;margin-bottom:16px}.graph-card img{width:100%;height:158mm;object-fit:contain}
      table{font-size:9px;table-layout:auto}th{background:#123a57!important;color:#fff;padding:9px 8px;border:1px solid #d8e3ea}td{padding:7px 8px;border:1px solid #dce6ec;line-height:1.45;overflow-wrap:anywhere}tbody tr:nth-child(even){background:#f5f9fc}.small th,.faculty-table th,.department-table th,.activity-table th{background:#123a57!important}
      thead{display:table-header-group}tr{break-inside:avoid}.report-section{page-break-after:always}
      @media print{button{display:none}}
    </style></head><body>${sectionHtmlParts.join("")}<script>window.onload=()=>{window.print();};<\/script></body></html>`);
  popup.document.close();
}
// Bind report buttons after the original dashboard DOM has been initialized.
window.addEventListener("DOMContentLoaded", () => {
  $("pdfBtn")?.addEventListener("click", generatePdfReport);
});


/* =======================================================
   SUBMISSION SECTIONS: directly read source workbook sheets
   ======================================================= */
function initSubmissionExplorer(workbook) {
  const sheet = name => workbook.Sheets[name]
    ? XLSX.utils.sheet_to_json(workbook.Sheets[name], {defval: "", raw: false}) : [];
  const activities = sheet("All Separated Activities");
  const mouRecords = sheet("MoU Country (No Nulls)").filter(r => clean(r.Country));
  const basicRecords = Array.from(new Map(activities.filter(r => clean(r["Ref #"]))
    .map(r => [r["Ref #"], r])).values());
  const ids = ["reportCampus","reportFaculty","reportDepartment","reportActivity","reportCountry","reportMouType"];
  const val = id => $(id)?.value || "All";
  const distinct = values => [...new Set(values.map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const options = (id,values,defaultName) => {
    const select = $(id); if(!select) return;
    const previous = select.value;
    select.replaceChildren();
    const make = (v,label) => {const o=document.createElement("option");o.value=v;o.textContent=label;select.appendChild(o)};
    make("All",defaultName);distinct(values).forEach(v=>make(v,v));
    select.value = [...select.options].some(o=>o.value===previous) ? previous : "All";
  };
  const matchBasic = r =>
    (val("reportCampus")==="All" || clean(r.Campus)===val("reportCampus")) &&
    (val("reportFaculty")==="All" || clean(r.Faculty)===val("reportFaculty")) &&
    (val("reportDepartment")==="All" || clean(r.Department)===val("reportDepartment"));
  const matchingFacultyDepartment = r =>
    (val("reportFaculty")==="All" || clean(r.Faculty)===val("reportFaculty")) &&
    (val("reportDepartment")==="All" || clean(r.Department)===val("reportDepartment"));
  const td = (row,text) => {const cell=document.createElement("td");cell.textContent=clean(text) || "—";row.appendChild(cell)};
  const table = (id,rows,fields,empty) => {
    const body=$(id);if(!body)return;body.replaceChildren();
    if(!rows.length){const tr=document.createElement("tr");const cell=document.createElement("td");cell.colSpan=fields.length;cell.textContent=empty;tr.appendChild(cell);body.appendChild(tr);return;}
    const fragment=document.createDocumentFragment();
    rows.forEach(r=>{const tr=document.createElement("tr");fields.forEach(k=>td(tr,r[k]));fragment.appendChild(tr)});
    body.appendChild(fragment);
  };
  const refresh = (changed) => {
    if(changed === "reportCampus"){$("reportFaculty").value="All";$("reportDepartment").value="All"}
    if(changed === "reportFaculty")$("reportDepartment").value="All";
    options("reportCampus",basicRecords.map(r=>r.Campus),"All Campuses");
    options("reportFaculty",basicRecords.filter(r=>val("reportCampus")==="All"||r.Campus===val("reportCampus")).map(r=>r.Faculty),"All Faculties");
    options("reportDepartment",basicRecords.filter(r=>(val("reportCampus")==="All"||r.Campus===val("reportCampus")) && (val("reportFaculty")==="All"||r.Faculty===val("reportFaculty"))).map(r=>r.Department),"All Departments");
    const basics=basicRecords.filter(matchBasic);
    const refs=new Set(basics.map(r=>clean(r["Ref #"])));
    if($("reportBasicSummary"))$("reportBasicSummary").textContent=`${basics.length} submissions • ${distinct(basics.map(r=>r.Campus)).length} campuses • ${distinct(basics.map(r=>r.Faculty)).length} faculties • ${distinct(basics.map(r=>r.Department)).length} departments`;
    const activityScope=activities.filter(r=>refs.has(clean(r["Ref #"])));
    options("reportActivity",activityScope.map(r=>clean(r["Separated Activity Type"])||r["Original Activity Type"]),"All Activity Types");
    const term=clean($("reportOrganization")?.value).toLowerCase();
    const shown=activityScope.filter(r=>(val("reportActivity")==="All"||(clean(r["Separated Activity Type"])||clean(r["Original Activity Type"]))===val("reportActivity")) && (!term||clean(r["Organization / Activity Title"]).toLowerCase().includes(term)));
    table("reportActivityBody",shown,["Faculty","Department","Separated Activity Type","How Many","Organization / Activity Title"],"No matching activities");
    const sum=shown.reduce((n,r)=>n+(Number(clean(r["How Many"]).replace(/,/g,""))||0),0);
    if($("reportActivityTotal"))$("reportActivityTotal").value=sum.toLocaleString();
    if($("reportActivityFooter"))$("reportActivityFooter").textContent=`${shown.length} activity records displayed • ${sum.toLocaleString()} total reported in “How Many”`;
    const mouScope=mouRecords.filter(r=>matchingFacultyDepartment(r) && (val("reportCampus")==="All" || refs.has(clean(r["Submission Ref"]))));
    options("reportCountry",mouScope.map(r=>r.Country),"All Countries");
    options("reportMouType",mouScope.map(r=>r["MoU Type"]),"All MoU Types");
    const partner=clean($("reportPartner")?.value).toLowerCase();
    const mous=mouScope.filter(r=>(val("reportCountry")==="All"||r.Country===val("reportCountry")) && (val("reportMouType")==="All"||r["MoU Type"]===val("reportMouType")) && (!partner||clean(r["Partner Institution"]).toLowerCase().includes(partner)));
    table("reportMouBody",mous,["Partner Institution","Country","MoU Type","Status","Faculty","Department"],"No matching MoU records with a valid country");
    if($("reportMouFooter"))$("reportMouFooter").textContent=`${mous.length} valid-country MoU entries displayed • Null and missing countries excluded • Evidence not included`;
  };
  ids.forEach(id=>$(id)?.addEventListener("change",()=>refresh(id)));
  ["reportOrganization","reportPartner"].forEach(id=>$(id)?.addEventListener("input",()=>refresh(id)));
  refresh();
}
