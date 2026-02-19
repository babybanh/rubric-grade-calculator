(() => {
  "use strict";

  const SCORE_OPTIONS = Array.from({ length: 11 }, (_, index) => index * 10);
  const WEIGHT_TOLERANCE = 0.01;
  const STORAGE_KEY = "rubric-grade-calculator-v2";

  const setupViewEl = document.getElementById("setupView");
  const graderViewEl = document.getElementById("graderView");
  const sectionCountEl = document.getElementById("sectionCount");
  const classCountEl = document.getElementById("classCount");
  const helpThresholdEl = document.getElementById("helpThreshold");
  const includeCommentsEl = document.getElementById("includeComments");
  const sectionsConfigEl = document.getElementById("sectionsConfig");
  const classesConfigEl = document.getElementById("classesConfig");
  const weightMessageEl = document.getElementById("weightMessage");
  const setupErrorEl = document.getElementById("setupError");
  const createSheetBtn = document.getElementById("createSheetBtn");
  const backToGraderBtn = document.getElementById("backToGraderBtn");
  const editSetupBtn = document.getElementById("editSetupBtn");
  const setupWarningDialogEl = document.getElementById("setupWarningDialog");
  const resetDataDialogEl = document.getElementById("resetDataDialog");
  const editSetupWrapEl = document.getElementById("editSetupWrap");
  const classesGradingEl = document.getElementById("classesGrading");
  const summaryTopEl = document.getElementById("summaryTop");
  const overallSummaryPanelEl = document.getElementById("overallSummaryPanel");
  const overallAverageEl = document.getElementById("overallAverage");
  const overallHelpCountEl = document.getElementById("overallHelpCount");
  const overallHelpListEl = document.getElementById("overallHelpList");
  const overallGradedCountEl = document.getElementById("overallGradedCount");
  const overallClassCountEl = document.getElementById("overallClassCount");
  const classSnapshotGridEl = document.getElementById("classSnapshotGrid");
  const overallNotesGridEl = document.getElementById("overallNotesGrid");
  const overallProgressNoteEl = document.getElementById("overallProgressNote");

  const globalSearchWrapEl = document.getElementById("globalSearchWrap");
  const globalSearchInputEl = document.getElementById("globalSearchInput");
  const globalSearchResultsEl = document.getElementById("globalSearchResults");
  const workspaceToolsEl = document.getElementById("workspaceTools");
  const setupImportBtn = document.getElementById("setupImportBtn");

  const saveBtn = document.getElementById("saveBtn");
  const resetEnteredDataBtn = document.getElementById("resetEnteredDataBtn");
  const exportBtn = document.getElementById("exportBtn");
  const importBtn = document.getElementById("importBtn");
  const importFileInputEl = document.getElementById("importFileInput");
  const saveStatusEl = document.getElementById("saveStatus");
  const lastSavedTextEl = document.getElementById("lastSavedText");
  const sheetExportClassSelectEl = document.getElementById("sheetExportClassSelect");
  const sheetExportStudentSelectEl = document.getElementById("sheetExportStudentSelect");
  const sheetExportModeSelectEl = document.getElementById("sheetExportModeSelect");
  const exportStudentSheetBtn = document.getElementById("exportStudentSheetBtn");
  const exportClassReportBtn = document.getElementById("exportClassReportBtn");

  const state = {
    setup: {
      sectionCount: 1,
      classCount: 1,
      helpThreshold: 70,
      includeComments: true,
      sections: [],
      classes: [],
    },
    grading: null,
    ui: {
      showAllHelp: false,
      dragSource: null,
      expandedSlots: {},
      studentSelectTypeahead: {
        classIndex: null,
        buffer: "",
        stamp: 0,
      },
      globalSearch: {
        activeIndex: -1,
        query: "",
      },
    },
  };

  let autoSaveTimer = null;
  let summaryCollapseTimer = null;
  let importRequestedFromSetup = false;

  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

  const numberOr = (value, fallback = 0) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  };

  const normalizeKey = (value) => String(value ?? "").trim().toLowerCase();
  const scoreOrNull = (value) => {
    if (value === null || value === undefined || String(value).trim() === "") return null;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? clamp(parsed, 0, 100) : null;
  };

  const stripDiacritics = (value) => {
    const raw = String(value ?? "");
    try {
      return raw.normalize("NFD").replace(/\p{Diacritic}/gu, "");
    } catch (error) {
      return raw;
    }
  };

  const normalizeSearchText = (value) => stripDiacritics(value).toLowerCase();

  const escapeRegExp = (value) => String(value ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  const htmlEscape = (value) =>
    String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  const cloneJson = (value) => JSON.parse(JSON.stringify(value));

  const formatStamp = (isoString) => {
    const parsedDate = new Date(isoString);
    if (Number.isNaN(parsedDate.getTime())) return "";
    return parsedDate.toLocaleString();
  };

  const setSaveStatus = (message, tone = "") => {
    if (!saveStatusEl) return;
    saveStatusEl.className = "save-status";
    if (tone) saveStatusEl.classList.add(tone);
    saveStatusEl.textContent = message;
  };

  const setLastSavedText = (isoString) => {
    if (!lastSavedTextEl) return;
    if (!isoString) {
      lastSavedTextEl.textContent = "Last saved: --";
      return;
    }
    const stamp = formatStamp(isoString);
    lastSavedTextEl.textContent = stamp
      ? `Last saved: ${stamp}`
      : "Last saved: --";
  };

  const enableNumberInputQuickReplace = (inputEl) => {
    if (!(inputEl instanceof HTMLInputElement)) return;
    const selectAll = () => {
      requestAnimationFrame(() => {
        try {
          inputEl.select();
        } catch (error) {
        }
      });
    };

    inputEl.addEventListener("focus", selectAll);
    inputEl.addEventListener("click", selectAll);
  };

  [sectionCountEl, classCountEl, helpThresholdEl].forEach(enableNumberInputQuickReplace);

  const highlightMatch = (text, query) => {
    const safeText = String(text ?? "");
    const trimmed = String(query ?? "").trim();
    if (!trimmed) return htmlEscape(safeText);

    try {
      const re = new RegExp(escapeRegExp(trimmed), "i");
      const match = safeText.match(re);
      if (!match || match.index === undefined) return htmlEscape(safeText);
      const start = match.index;
      const end = start + match[0].length;
      return `${htmlEscape(safeText.slice(0, start))}<mark>${htmlEscape(
        safeText.slice(start, end)
      )}</mark>${htmlEscape(safeText.slice(end))}`;
    } catch (error) {
      return htmlEscape(safeText);
    }
  };

  const scoreSearchMatch = (haystackRaw, queryRaw) => {
    const query = normalizeSearchText(queryRaw).trim();
    if (!query) return null;
    const haystack = normalizeSearchText(haystackRaw);
    const index = haystack.indexOf(query);
    if (index < 0) return null;

    let score = 0;
    if (index == 0) score += 120;
    if (index > 0 && /\s/.test(haystack[index - 1])) score += 80;
    score += Math.max(0, 40 - index);
    score += Math.min(40, query.length * 4);
    return { score, index };
  };

  const buildSearchCandidates = () => {
    if (!state.grading) return [];
    const candidates = [];

    state.grading.classes.forEach((classRecord, classIndex) => {
      candidates.push({
        kind: "class",
        classIndex,
        label: classRecord.name,
        secondary: "Class",
        haystack: `${classRecord.name}`,
      });

      const classNote = String(classRecord.classSupportNote ?? "").trim();
      if (classNote) {
        candidates.push({
          kind: "class-note",
          classIndex,
          label: classRecord.name,
          secondary: "Class note",
          haystack: `${classRecord.name} ${classNote}`,
          snippet: classNote,
        });
      }

      classRecord.students.forEach((studentRecord, studentIndex) => {
        candidates.push({
          kind: "student",
          classIndex,
          studentIndex,
          label: studentRecord.name,
          secondary: classRecord.name,
          haystack: `${studentRecord.name} ${classRecord.name}`,
        });

        studentRecord.sections.forEach((sectionRecord, sectionIndex) => {
          const comment = String(sectionRecord.comment ?? "").trim();
          if (!comment) return;
          const sectionName = state.grading.sections?.[sectionIndex]?.name ?? `Section ${sectionIndex + 1}`;
          candidates.push({
            kind: "section-comment",
            classIndex,
            studentIndex,
            sectionIndex,
            label: studentRecord.name,
            secondary: `${classRecord.name} · ${sectionName} comment`,
            haystack: `${studentRecord.name} ${classRecord.name} ${sectionName} ${comment}`,
            snippet: comment,
          });
        });
      });
    });

    const overallNote = String(state.grading.overallProgressNote ?? "").trim();
    if (overallNote) {
      candidates.push({
        kind: "overall-note",
        label: "Overall notes",
        secondary: "Overall",
        haystack: `overall notes ${overallNote}`,
        snippet: overallNote,
      });
    }

    return candidates;
  };

  const buildSnippet = (text, query, limit = 96) => {
    const raw = String(text ?? "").replace(/\s+/g, " ").trim();
    if (!raw) return "";
    const q = normalizeSearchText(query).trim();
    if (!q) return raw.length > limit ? `${raw.slice(0, limit - 1)}…` : raw;

    const normalized = normalizeSearchText(raw);
    const idx = normalized.indexOf(q);
    if (idx < 0) return raw.length > limit ? `${raw.slice(0, limit - 1)}…` : raw;

    const start = Math.max(0, idx - Math.floor(limit / 2));
    const end = Math.min(raw.length, start + limit);
    const prefix = start > 0 ? "…" : "";
    const suffix = end < raw.length ? "…" : "";
    return `${prefix}${raw.slice(start, end)}${suffix}`;
  };

  const closeGlobalSearch = () => {
    if (!globalSearchResultsEl) return;
    globalSearchResultsEl.classList.add("hidden");
    globalSearchResultsEl.innerHTML = "";
    state.ui.globalSearch.activeIndex = -1;
  };

  const openGlobalSearchResults = (html) => {
    if (!globalSearchResultsEl) return;
    globalSearchResultsEl.innerHTML = html;
    globalSearchResultsEl.classList.remove("hidden");
  };

  const setGlobalSearchActiveIndex = (nextIndex) => {
    state.ui.globalSearch.activeIndex = nextIndex;
    const items = globalSearchResultsEl?.querySelectorAll("[data-role='search-hit']") ?? [];
    items.forEach((item, index) => {
      if (!(item instanceof HTMLElement)) return;
      item.classList.toggle("active", index === nextIndex);
      item.setAttribute("aria-selected", index === nextIndex ? "true" : "false");
    });
  };

  const focusAndFlash = (el) => {
    if (!(el instanceof HTMLElement)) return;
    try {
      el.scrollIntoView({ block: "center", behavior: "smooth" });
    } catch (error) {
      el.scrollIntoView();
    }
    el.classList.add("search-flash");
    setTimeout(() => el.classList.remove("search-flash"), 1100);
    if (el instanceof HTMLTextAreaElement || el instanceof HTMLInputElement) {
      el.focus();
    }
  };

  const applySearchSelection = (hit) => {
    if (!state.grading || !hit) return;

    const setActiveClass = (classIndex) => {
      if (!Number.isInteger(classIndex)) return;
      if (classIndex < 0 || classIndex >= state.grading.classes.length) return;
      state.grading.activeClassIndex = classIndex;
      if (sheetExportClassSelectEl) sheetExportClassSelectEl.value = String(classIndex);
    };

    if (hit.kind === "class") {
      setActiveClass(hit.classIndex);
      collapseOverallSummary();
      renderGrader();
      scheduleAutoSave();
      focusAndFlash(classesGradingEl.querySelector(".class-card"));
      return;
    }

    if (hit.kind === "class-note") {
      setActiveClass(hit.classIndex);
      collapseOverallSummary();
      renderGrader();
      scheduleAutoSave();
      requestAnimationFrame(() => {
        const note = classesGradingEl.querySelector(
          `textarea[data-type="class-support-note"][data-class-index="${hit.classIndex}"]`
        );
        if (note) focusAndFlash(note);
      });
      return;
    }

    if (hit.kind === "student" || hit.kind === "section-comment") {
      setActiveClass(hit.classIndex);
      const classRecord = state.grading.classes[hit.classIndex];
      if (!classRecord) return;
      classRecord.selectedStudentIndex = hit.studentIndex;
      collapseOverallSummary();
      renderGrader();
      scheduleAutoSave();

      if (sheetExportStudentSelectEl && !sheetExportStudentSelectEl.disabled) {
        sheetExportStudentSelectEl.value = String(hit.studentIndex);
      }

      requestAnimationFrame(() => {
        if (hit.kind === "student") {
          focusAndFlash(classesGradingEl.querySelector(".student-name"));
          return;
        }

        const sectionCard = classesGradingEl.querySelector(
          `[data-type="section-card"][data-class-index="${hit.classIndex}"][data-student-index="${hit.studentIndex}"][data-section-index="${hit.sectionIndex}"]`
        );
        if (!sectionCard) return;
        focusAndFlash(sectionCard);
        const textarea = sectionCard.querySelector("textarea");
        if (textarea) focusAndFlash(textarea);
      });
      return;
    }

    if (hit.kind === "overall-note") {
      if (overallSummaryPanelEl) {
        overallSummaryPanelEl.open = true;
        clearTimeout(summaryCollapseTimer);
        overallSummaryPanelEl.classList.remove("closing");
      }
      renderGrader();
      scheduleAutoSave();
      requestAnimationFrame(() => {
        if (overallProgressNoteEl) focusAndFlash(overallProgressNoteEl);
      });
    }
  };

  const renderGlobalSearch = (queryRaw) => {
    const query = String(queryRaw ?? "");
    const trimmed = query.trim();
    state.ui.globalSearch.query = query;

    if (!globalSearchResultsEl) return;
    if (!state.grading || !trimmed) {
      closeGlobalSearch();
      return;
    }

    const candidates = buildSearchCandidates();
    const hits = [];

    candidates.forEach((candidate) => {
      const match = scoreSearchMatch(candidate.haystack, trimmed);
      if (!match) return;
      hits.push({ ...candidate, score: match.score });
    });

    hits.sort((a, b) => b.score - a.score || a.label.localeCompare(b.label));
    const topHits = hits.slice(0, 10);

    if (topHits.length === 0) {
      openGlobalSearchResults('<div class="search-empty">No matches.</div>');
      setGlobalSearchActiveIndex(-1);
      return;
    }

    const html = topHits
      .map((hit, index) => {
        const snippetSource = hit.snippet ? buildSnippet(hit.snippet, trimmed) : "";
        const primaryHtml = highlightMatch(hit.label, trimmed);
        const secondaryHtml = htmlEscape(hit.secondary || "");
        const snippetHtml = snippetSource
          ? `<div class="search-snippet">${highlightMatch(snippetSource, trimmed)}</div>`
          : "";

        const payload = encodeURIComponent(
          JSON.stringify({
            kind: hit.kind,
            classIndex: Number.isInteger(hit.classIndex) ? hit.classIndex : null,
            studentIndex: Number.isInteger(hit.studentIndex) ? hit.studentIndex : null,
            sectionIndex: Number.isInteger(hit.sectionIndex) ? hit.sectionIndex : null,
          })
        );

        return `
          <button
            type="button"
            class="search-hit"
            data-role="search-hit"
            data-payload="${payload}"
            role="option"
            aria-selected="${index === 0 ? "true" : "false"}"
          >
            <div class="search-hit-top">
              <div class="search-hit-primary">${primaryHtml}</div>
              <div class="search-hit-secondary">${secondaryHtml}</div>
            </div>
            ${snippetHtml}
          </button>
        `;
      })
      .join("");

    openGlobalSearchResults(html);
    setGlobalSearchActiveIndex(0);
  };

  const showSetupView = () => {
    graderViewEl.classList.add("hidden");
    setupViewEl.classList.remove("hidden");
    editSetupWrapEl.classList.add("hidden");
    if (workspaceToolsEl) workspaceToolsEl.classList.add("hidden");
    if (globalSearchWrapEl) globalSearchWrapEl.classList.add("hidden");
    if (globalSearchInputEl) globalSearchInputEl.value = "";
    closeGlobalSearch();
  };

  const showGraderView = () => {
    setupViewEl.classList.add("hidden");
    graderViewEl.classList.remove("hidden");
    editSetupWrapEl.classList.remove("hidden");
    if (workspaceToolsEl) workspaceToolsEl.classList.remove("hidden");
    if (globalSearchWrapEl) globalSearchWrapEl.classList.remove("hidden");
  };

  const collapseOverallSummary = () => {
    if (!overallSummaryPanelEl || !overallSummaryPanelEl.open) return;
    overallSummaryPanelEl.classList.add("closing");
    clearTimeout(summaryCollapseTimer);
    summaryCollapseTimer = setTimeout(() => {
      if (!overallSummaryPanelEl) return;
      overallSummaryPanelEl.open = false;
      overallSummaryPanelEl.classList.remove("closing");
    }, 220);
  };

  const confirmSetupWarning = () =>
    new Promise((resolve) => {
      if (!(setupWarningDialogEl instanceof HTMLDialogElement)) {
        resolve(
          window.confirm(
            "Returning to setup is safe. Applying setup changes keeps matching grades where possible, but removed classes/students/sections will be dropped. Continue?"
          )
        );
        return;
      }

      setupWarningDialogEl.returnValue = "";
      const onClose = () => {
        resolve(setupWarningDialogEl.returnValue === "continue");
      };
      setupWarningDialogEl.addEventListener("close", onClose, { once: true });
      setupWarningDialogEl.showModal();
    });

  const confirmResetDataWarning = () =>
    new Promise((resolve) => {
      if (!(resetDataDialogEl instanceof HTMLDialogElement)) {
        resolve(
          window.confirm(
            "Reset entered scores? This clears all grades, overrides, and notes, but keeps setup, classes, students, and section names."
          )
        );
        return;
      }

      resetDataDialogEl.returnValue = "";
      const onClose = () => {
        resolve(resetDataDialogEl.returnValue === "continue");
      };
      resetDataDialogEl.addEventListener("close", onClose, { once: true });
      resetDataDialogEl.showModal();
    });

  const createDefaultSection = (index) => ({
    name: `Section ${index + 1}`,
    weight: index === 0 ? 100 : 0,
    slots: 2,
    scoringMode: "average",
    allowDeductions: false,
    itemNames: ["Item 1", "Item 2"],
  });

  const createDefaultClass = (index) => ({
    name: `Class ${index + 1}`,
    studentCount: 5,
    studentNamesText: "",
  });

  const ensureSetupArrays = () => {
    if (!Array.isArray(state.setup.sections)) state.setup.sections = [];
    if (!Array.isArray(state.setup.classes)) state.setup.classes = [];

    const sectionCount = clamp(Math.round(numberOr(state.setup.sectionCount, 1)), 1, 12);
    const classCount = clamp(Math.round(numberOr(state.setup.classCount, 1)), 1, 12);

    state.setup.sectionCount = sectionCount;
    state.setup.classCount = classCount;

    while (state.setup.sections.length < sectionCount) {
      state.setup.sections.push(createDefaultSection(state.setup.sections.length));
    }
    if (state.setup.sections.length > sectionCount) {
      state.setup.sections.length = sectionCount;
    }

    while (state.setup.classes.length < classCount) {
      state.setup.classes.push(createDefaultClass(state.setup.classes.length));
    }
    if (state.setup.classes.length > classCount) {
      state.setup.classes.length = classCount;
    }

    state.setup.sections.forEach((section, index) => {
      if (!section.name) section.name = `Section ${index + 1}`;
      section.weight = clamp(numberOr(section.weight, 0), 0, 100);
      section.slots = clamp(Math.round(numberOr(section.slots, 1)), 1, 12);
      section.scoringMode = section.scoringMode === "highest" ? "highest" : "average";
      section.allowDeductions = Boolean(section.allowDeductions);
      if (!Array.isArray(section.itemNames)) section.itemNames = [];
      section.itemNames = Array.from({ length: section.slots }, (_, slotIndex) => {
        const rawName = section.itemNames[slotIndex];
        const fallback = `Item ${slotIndex + 1}`;
        return String(rawName ?? fallback).trim() || fallback;
      });
    });

    state.setup.classes.forEach((classItem, index) => {
      if (!classItem.name) classItem.name = `Class ${index + 1}`;
      classItem.studentCount = clamp(Math.round(numberOr(classItem.studentCount, 1)), 1, 60);
      if (typeof classItem.studentNamesText !== "string") classItem.studentNamesText = "";
    });
  };

  const sanitizeSectionConfig = (section, index, fallbackSlots = 2) => {
    const slots = clamp(Math.round(numberOr(section?.slots, fallbackSlots)), 1, 12);
    const sourceNames = Array.isArray(section?.itemNames) ? section.itemNames : [];
    const itemNames = Array.from({ length: slots }, (_, slotIndex) => {
      const fallback = `Item ${slotIndex + 1}`;
      return String(sourceNames[slotIndex] ?? fallback).trim() || fallback;
    });

    return {
      name: String(section?.name ?? `Section ${index + 1}`).trim() || `Section ${index + 1}`,
      weight: clamp(numberOr(section?.weight, 0), 0, 100),
      slots,
      scoringMode: section?.scoringMode === "highest" ? "highest" : "average",
      allowDeductions: Boolean(section?.allowDeductions),
      itemNames,
    };
  };

  const normalizeSetup = (rawSetup) => {
    const nextSetup = {
      sectionCount: clamp(Math.round(numberOr(rawSetup?.sectionCount, 1)), 1, 12),
      classCount: clamp(Math.round(numberOr(rawSetup?.classCount, 1)), 1, 12),
      helpThreshold: clamp(Math.round(numberOr(rawSetup?.helpThreshold, 70)), 0, 100),
      includeComments: rawSetup?.includeComments !== false,
      sections: Array.isArray(rawSetup?.sections) ? rawSetup.sections : [],
      classes: Array.isArray(rawSetup?.classes) ? rawSetup.classes : [],
    };

    nextSetup.sections = nextSetup.sections.map((section, index) =>
      sanitizeSectionConfig(section, index)
    );

    nextSetup.classes = nextSetup.classes.map((classItem, index) => ({
      name: String(classItem?.name ?? `Class ${index + 1}`).trim() || `Class ${index + 1}`,
      studentCount: clamp(Math.round(numberOr(classItem?.studentCount, 5)), 1, 60),
      studentNamesText: String(classItem?.studentNamesText ?? ""),
    }));

    state.setup = nextSetup;
    ensureSetupArrays();
  };

  const normalizeGrading = (rawGrading) => {
    if (!rawGrading || typeof rawGrading !== "object") return null;

    const sourceSections =
      Array.isArray(rawGrading.sections) && rawGrading.sections.length > 0
        ? rawGrading.sections
        : state.setup.sections;

    const sections = sourceSections.map((section, index) => sanitizeSectionConfig(section, index));

    const includeComments =
      typeof rawGrading.includeComments === "boolean"
        ? rawGrading.includeComments
        : state.setup.includeComments;

    const helpThreshold = clamp(
      Math.round(numberOr(rawGrading.helpThreshold, state.setup.helpThreshold)),
      0,
      100
    );

    const rawClasses = Array.isArray(rawGrading.classes) ? rawGrading.classes : [];

    const classes = rawClasses.map((classItem, classIndex) => {
      const rawStudents = Array.isArray(classItem.students) ? classItem.students : [];
      const students = rawStudents.map((student, studentIndex) => {
        const rawStudentSections = Array.isArray(student.sections) ? student.sections : [];

        const normalizedSections = sections.map((sectionConfig, sectionIndex) => {
          const rawStudentSection = rawStudentSections[sectionIndex] || {};

          const scores = Array.from({ length: sectionConfig.slots }, (_, slotIndex) => {
            return scoreOrNull(rawStudentSection.scores?.[slotIndex]);
          });

          const deductions = Array.from({ length: sectionConfig.slots }, (_, slotIndex) => {
            const value = rawStudentSection.deductions?.[slotIndex];
            if (!Number.isFinite(Number(value))) return 0;
            return clamp(Math.round(Number(value)), 0, 100);
          });

          return {
            scores,
            deductions,
            overrideScore: scoreOrNull(rawStudentSection.overrideScore),
            comment: String(rawStudentSection.comment ?? ""),
          };
        });

        return {
          name: String(student?.name ?? `Student ${studentIndex + 1}`).trim() || `Student ${studentIndex + 1}`,
          totalOverride: scoreOrNull(student?.totalOverride),
          sections: normalizedSections,
        };
      });

      const selectedStudentIndex = Number(classItem.selectedStudentIndex);
      const hasValidSelectedIndex =
        Number.isInteger(selectedStudentIndex) &&
        selectedStudentIndex >= 0 &&
        selectedStudentIndex < students.length;

      return {
        name: String(classItem?.name ?? `Class ${classIndex + 1}`).trim() || `Class ${classIndex + 1}`,
        classSupportNote: String(classItem.classSupportNote ?? ""),
        sectionOrder: (() => {
          const fromClass = Array.isArray(classItem?.sectionOrder) ? classItem.sectionOrder : [];
          const fromLegacyStudent = Array.isArray(classItem?.students?.[0]?.sectionOrder)
            ? classItem.students[0].sectionOrder
            : [];
          const rawOrder = fromClass.length > 0 ? fromClass : fromLegacyStudent;
          const uniqueValid = [];
          rawOrder.forEach((value) => {
            const index = Number(value);
            if (!Number.isInteger(index)) return;
            if (index < 0 || index >= sections.length) return;
            if (uniqueValid.includes(index)) return;
            uniqueValid.push(index);
          });
          for (let index = 0; index < sections.length; index += 1) {
            if (!uniqueValid.includes(index)) uniqueValid.push(index);
          }
          return uniqueValid;
        })(),
        selectedStudentIndex: hasValidSelectedIndex ? selectedStudentIndex : null,
        students,
      };
    });

    const rawActiveClassIndex = Number(rawGrading.activeClassIndex);
    const activeClassIndex =
      Number.isInteger(rawActiveClassIndex) &&
      rawActiveClassIndex >= 0 &&
      rawActiveClassIndex < classes.length
        ? rawActiveClassIndex
        : null;

    return {
      sections,
      includeComments,
      helpThreshold,
      overallProgressNote: String(rawGrading.overallProgressNote ?? ""),
      activeClassIndex,
      classes,
    };
  };

  const totalWeight = () =>
    state.setup.sections.reduce((sum, section) => sum + numberOr(section.weight, 0), 0);

  const weightStatusMessage = () => {
    const total = totalWeight();
    const delta = Math.abs(total - 100);

    if (delta <= WEIGHT_TOLERANCE) {
      return {
        text: `Total section weight: ${total.toFixed(2)}% (ready)`,
        className: "small-note good",
      };
    }

    return {
      text: `Total section weight: ${total.toFixed(2)}% (must equal 100%)`,
      className: "small-note bad",
    };
  };

  const renderSectionEditors = () => {
    sectionsConfigEl.innerHTML = state.setup.sections
      .map(
        (section, sectionIndex) => `
        <article class="config-block">
          <div class="config-title">Section ${sectionIndex + 1}</div>
          <div class="grid grid-3">
            <label>
              Name
              <input
                type="text"
                data-kind="section"
                data-field="name"
                data-section-index="${sectionIndex}"
                value="${htmlEscape(section.name)}"
              />
            </label>
            <label>
              Weight (%)
              <input
                type="number"
                min="0"
                max="100"
                step="0.1"
                data-kind="section"
                data-field="weight"
                data-section-index="${sectionIndex}"
                value="${section.weight}"
              />
            </label>
            <label>
              Number of grade slots
              <input
                type="number"
                min="1"
                max="12"
                step="1"
                data-kind="section"
                data-field="slots"
                data-section-index="${sectionIndex}"
                value="${section.slots}"
              />
            </label>
            <label>
              Slot scoring
              <select
                data-kind="section"
                data-field="scoringMode"
                data-section-index="${sectionIndex}"
              >
                <option value="average" ${
                  section.scoringMode === "average" ? "selected" : ""
                }>
                  Average entered scores
                </option>
                <option value="highest" ${
                  section.scoringMode === "highest" ? "selected" : ""
                }>
                  Highest score only
                </option>
              </select>
            </label>
            <label class="checkbox-label">
              <input
                type="checkbox"
                data-kind="section"
                data-field="allowDeductions"
                data-section-index="${sectionIndex}"
                ${section.allowDeductions ? "checked" : ""}
              />
              Add deduction field per slot
            </label>
          </div>
        </article>
      `
      )
      .join("");
  };

  const renderClassEditors = () => {
    classesConfigEl.innerHTML = state.setup.classes
      .map(
        (classItem, classIndex) => `
        <article class="config-block">
          <div class="config-title">Class ${classIndex + 1}</div>
          <div class="grid grid-2">
            <label>
              Class name
              <input
                type="text"
                data-kind="class"
                data-field="name"
                data-class-index="${classIndex}"
                value="${htmlEscape(classItem.name)}"
              />
            </label>
            <label>
              Number of students
              <input
                type="number"
                min="1"
                max="60"
                step="1"
                data-kind="class"
                data-field="studentCount"
                data-class-index="${classIndex}"
                value="${classItem.studentCount}"
              />
            </label>
          </div>
          <label>
            Student names (one per line, optional)
            <textarea
              rows="3"
              data-kind="class"
              data-field="studentNamesText"
              data-class-index="${classIndex}"
              placeholder="Leave blank to auto-create Student 1, Student 2, ...">${htmlEscape(
                classItem.studentNamesText
              )}</textarea>
          </label>
        </article>
      `
      )
      .join("");
  };

  const updateWeightMessage = () => {
    const status = weightStatusMessage();
    weightMessageEl.className = status.className;
    weightMessageEl.textContent = status.text;
  };

  const renderSetup = () => {
    ensureSetupArrays();
    sectionCountEl.value = String(state.setup.sectionCount);
    classCountEl.value = String(state.setup.classCount);
    helpThresholdEl.value = String(state.setup.helpThreshold);
    includeCommentsEl.checked = state.setup.includeComments;
    if (backToGraderBtn) backToGraderBtn.classList.toggle("hidden", !state.grading);
    if (createSheetBtn) {
      createSheetBtn.textContent = state.grading ? "Apply Setup Changes" : "Create Grading Sheet";
    }
    renderSectionEditors();
    renderClassEditors();
    updateWeightMessage();
  };

  const parseStudentNames = (classConfig) => {
    const enteredNames = classConfig.studentNamesText
      .split(/\n/g)
      .map((name) => name.trim())
      .filter(Boolean);

    return Array.from({ length: classConfig.studentCount }, (_, index) => {
      return enteredNames[index] || `Student ${index + 1}`;
    });
  };

  const buildSetupFromGrading = (gradingState) => ({
    sectionCount: gradingState.sections.length,
    classCount: gradingState.classes.length,
    helpThreshold: clamp(Math.round(numberOr(gradingState.helpThreshold, 70)), 0, 100),
    includeComments: gradingState.includeComments !== false,
    sections: gradingState.sections.map((section, sectionIndex) =>
      sanitizeSectionConfig(section, sectionIndex)
    ),
    classes: gradingState.classes.map((classRecord, classIndex) => ({
      name: String(classRecord?.name ?? `Class ${classIndex + 1}`).trim() || `Class ${classIndex + 1}`,
      studentCount: Math.max(1, classRecord?.students?.length || 1),
      studentNamesText: Array.isArray(classRecord?.students)
        ? classRecord.students.map((student, studentIndex) => {
            const fallback = `Student ${studentIndex + 1}`;
            return String(student?.name ?? fallback).trim() || fallback;
          }).join("\n")
        : "",
    })),
  });

  const syncSetupFromGrading = () => {
    if (!state.grading) return;
    state.setup = buildSetupFromGrading(state.grading);
    ensureSetupArrays();
  };

  const validateSetup = () => {
    const errors = [];

    if (Math.abs(totalWeight() - 100) > WEIGHT_TOLERANCE) {
      errors.push("Section weights must add up to 100%.");
    }

    state.setup.sections.forEach((section, index) => {
      if (!section.name.trim()) errors.push(`Section ${index + 1} needs a name.`);
      if (section.weight < 0 || section.weight > 100) {
        errors.push(`Section \"${section.name || `#${index + 1}`}\" has an invalid weight.`);
      }
      if (section.slots < 1 || section.slots > 12) {
        errors.push(`Section \"${section.name || `#${index + 1}`}\" needs 1 to 12 grade slots.`);
      }
    });

    state.setup.classes.forEach((classItem, index) => {
      if (classItem.studentCount < 1 || classItem.studentCount > 60) {
        errors.push(`Class ${index + 1} needs 1 to 60 students.`);
      }
    });

    return errors;
  };

  const buildGradingState = () => {
    const sections = state.setup.sections.map((section, index) => sanitizeSectionConfig(section, index));

    return {
      sections,
      includeComments: state.setup.includeComments,
      helpThreshold: state.setup.helpThreshold,
      overallProgressNote: "",
      activeClassIndex: null,
      classes: state.setup.classes.map((classItem, classIndex) => ({
        name: classItem.name.trim() || `Class ${classIndex + 1}`,
        classSupportNote: "",
        sectionOrder: sections.map((_, sectionIndex) => sectionIndex),
        selectedStudentIndex: null,
        students: parseStudentNames(classItem).map((studentName) => ({
          name: studentName,
          totalOverride: null,
          sections: sections.map((section) => ({
            scores: Array(section.slots).fill(null),
            deductions: Array(section.slots).fill(0),
            overrideScore: null,
            comment: "",
          })),
        })),
      })),
    };
  };

  const createBlankStudentRecord = (name) => ({
    name,
    totalOverride: null,
    sections: state.grading.sections.map((sectionConfig) => ({
      scores: Array(sectionConfig.slots).fill(null),
      deductions: Array(sectionConfig.slots).fill(0),
      overrideScore: null,
      comment: "",
    })),
  });

  const nextClassName = () => {
    const existing = new Set(state.grading.classes.map((classItem) => normalizeKey(classItem.name)));
    for (let index = 1; index <= 999; index += 1) {
      const candidate = `Class ${index}`;
      if (!existing.has(normalizeKey(candidate))) return candidate;
    }
    return `Class ${state.grading.classes.length + 1}`;
  };

  const nextStudentName = (classRecord) => {
    const existing = new Set(classRecord.students.map((student) => normalizeKey(student.name)));
    for (let index = 1; index <= 999; index += 1) {
      const candidate = `Student ${index}`;
      if (!existing.has(normalizeKey(candidate))) return candidate;
    }
    return `Student ${classRecord.students.length + 1}`;
  };

  const clearEnteredGradingData = () => {
    if (!state.grading) return;

    state.grading.overallProgressNote = "";
    state.ui.showAllHelp = false;
    state.ui.expandedSlots = {};

    state.grading.classes.forEach((classRecord) => {
      classRecord.classSupportNote = "";

      classRecord.students.forEach((studentRecord) => {
        studentRecord.totalOverride = null;
        studentRecord.sections.forEach((sectionRecord) => {
          sectionRecord.overrideScore = null;
          sectionRecord.comment = "";
          sectionRecord.scores = sectionRecord.scores.map(() => null);
          sectionRecord.deductions = sectionRecord.deductions.map(() => 0);
        });
      });
    });
  };

  const addSectionItemSlot = (sectionIndex) => {
    const sectionConfig = state.grading?.sections?.[sectionIndex];
    if (!sectionConfig) return false;
    if (sectionConfig.slots >= 12) return false;

    sectionConfig.slots += 1;
    sectionConfig.itemNames.push(`Item ${sectionConfig.slots}`);

    if (state.setup.sections[sectionIndex]) {
      state.setup.sections[sectionIndex].slots = sectionConfig.slots;
      state.setup.sections[sectionIndex].itemNames = [...sectionConfig.itemNames];
    }

    state.grading.classes.forEach((classRecord) => {
      classRecord.students.forEach((studentRecord) => {
        const studentSection = studentRecord.sections[sectionIndex];
        if (!studentSection) return;
        studentSection.scores.push(null);
        studentSection.deductions.push(0);
      });
    });

    return true;
  };

  const removeSectionItemSlot = (sectionIndex) => {
    const sectionConfig = state.grading?.sections?.[sectionIndex];
    if (!sectionConfig) return { ok: false, reason: "missing" };
    if (sectionConfig.slots <= 1) return { ok: false, reason: "min" };

    const removeIndex = sectionConfig.slots - 1;
    const hasDataToDrop = state.grading.classes.some((classRecord) =>
      classRecord.students.some((studentRecord) => {
        const studentSection = studentRecord.sections[sectionIndex];
        if (!studentSection) return false;
        const score = studentSection.scores[removeIndex];
        const deduction = Number(studentSection.deductions[removeIndex]);
        return Number.isFinite(score) || (Number.isFinite(deduction) && deduction !== 0);
      })
    );

    if (hasDataToDrop) {
      const confirmed = window.confirm(
        "Remove last item in this section? Existing scores in that last item will be deleted."
      );
      if (!confirmed) return { ok: false, reason: "cancelled" };
    }

    sectionConfig.slots -= 1;
    sectionConfig.itemNames.pop();

    if (state.setup.sections[sectionIndex]) {
      state.setup.sections[sectionIndex].slots = sectionConfig.slots;
      state.setup.sections[sectionIndex].itemNames = [...sectionConfig.itemNames];
    }

    state.grading.classes.forEach((classRecord) => {
      classRecord.students.forEach((studentRecord) => {
        const studentSection = studentRecord.sections[sectionIndex];
        if (!studentSection) return;
        studentSection.scores.length = sectionConfig.slots;
        studentSection.deductions.length = sectionConfig.slots;
      });
    });

    return { ok: true, reason: "" };
  };

  const mergeExistingGradingWithSetup = (existingGrading) => {
    const sections = state.setup.sections.map((section, index) => sanitizeSectionConfig(section, index));
    const oldSections = Array.isArray(existingGrading?.sections) ? existingGrading.sections : [];

    const usedOldSectionIndexes = new Set();
    const sectionMatch = sections.map((sectionConfig, sectionIndex) => {
      const nameKey = normalizeKey(sectionConfig.name);
      let oldSectionIndex = oldSections.findIndex(
        (oldSection, oldIndex) =>
          !usedOldSectionIndexes.has(oldIndex) && normalizeKey(oldSection?.name) === nameKey
      );

      if (
        oldSectionIndex < 0 &&
        sectionIndex < oldSections.length &&
        !usedOldSectionIndexes.has(sectionIndex)
      ) {
        oldSectionIndex = sectionIndex;
      }

      if (oldSectionIndex >= 0) usedOldSectionIndexes.add(oldSectionIndex);
      return oldSectionIndex;
    });

    const oldToNewSectionIndex = {};
    sectionMatch.forEach((oldIndex, newIndex) => {
      if (oldIndex >= 0) oldToNewSectionIndex[oldIndex] = newIndex;
    });

    const oldClasses = Array.isArray(existingGrading?.classes) ? existingGrading.classes : [];
    const usedOldClassIndexes = new Set();

    const classes = state.setup.classes.map((classConfig, classIndex) => {
      const className = String(classConfig.name ?? `Class ${classIndex + 1}`).trim() || `Class ${classIndex + 1}`;
      const classKey = normalizeKey(className);

      let oldClassIndex = oldClasses.findIndex(
        (oldClass, oldIndex) => !usedOldClassIndexes.has(oldIndex) && normalizeKey(oldClass?.name) === classKey
      );
      if (oldClassIndex < 0 && classIndex < oldClasses.length && !usedOldClassIndexes.has(classIndex)) {
        oldClassIndex = classIndex;
      }

      const oldClass = oldClassIndex >= 0 ? oldClasses[oldClassIndex] : null;
      if (oldClassIndex >= 0) usedOldClassIndexes.add(oldClassIndex);

      const targetStudentNames = parseStudentNames(classConfig);
      const oldStudents = Array.isArray(oldClass?.students) ? oldClass.students : [];
      const usedOldStudentIndexes = new Set();

      const students = targetStudentNames.map((studentName, studentIndex) => {
        const studentKey = normalizeKey(studentName);
        let oldStudentIndex = oldStudents.findIndex(
          (oldStudent, index) =>
            !usedOldStudentIndexes.has(index) && normalizeKey(oldStudent?.name) === studentKey
        );

        if (
          oldStudentIndex < 0 &&
          studentIndex < oldStudents.length &&
          !usedOldStudentIndexes.has(studentIndex)
        ) {
          oldStudentIndex = studentIndex;
        }

        const oldStudent = oldStudentIndex >= 0 ? oldStudents[oldStudentIndex] : null;
        if (oldStudentIndex >= 0) usedOldStudentIndexes.add(oldStudentIndex);

        return {
          name: studentName,
          totalOverride: scoreOrNull(oldStudent?.totalOverride),
          sections: sections.map((sectionConfig, newSectionIndex) => {
            const oldSectionIndex = sectionMatch[newSectionIndex];
            const oldSection =
              oldStudent &&
              oldSectionIndex >= 0 &&
              Array.isArray(oldStudent.sections)
                ? oldStudent.sections[oldSectionIndex]
                : null;

            return {
              scores: Array.from({ length: sectionConfig.slots }, (_, slotIndex) => {
                return scoreOrNull(oldSection?.scores?.[slotIndex]);
              }),
              deductions: Array.from({ length: sectionConfig.slots }, (_, slotIndex) => {
                const deductionValue = oldSection?.deductions?.[slotIndex];
                if (!Number.isFinite(Number(deductionValue))) return 0;
                return clamp(Math.round(Number(deductionValue)), 0, 100);
              }),
              overrideScore: scoreOrNull(oldSection?.overrideScore),
              comment: String(oldSection?.comment ?? ""),
            };
          }),
        };
      });

      const sectionOrder = (() => {
        const rawOrder = Array.isArray(oldClass?.sectionOrder) ? oldClass.sectionOrder : [];
        const mapped = [];
        rawOrder.forEach((oldIndex) => {
          const newIndex = oldToNewSectionIndex[Number(oldIndex)];
          if (!Number.isInteger(newIndex)) return;
          if (newIndex < 0 || newIndex >= sections.length) return;
          if (mapped.includes(newIndex)) return;
          mapped.push(newIndex);
        });
        for (let index = 0; index < sections.length; index += 1) {
          if (!mapped.includes(index)) mapped.push(index);
        }
        return mapped;
      })();

      const selectedStudentIndex = (() => {
        const oldIndex = Number(oldClass?.selectedStudentIndex);
        if (!Number.isInteger(oldIndex) || oldIndex < 0 || oldIndex >= oldStudents.length) return null;
        const oldSelectedName = normalizeKey(oldStudents[oldIndex]?.name);
        const nextIndex = students.findIndex((student) => normalizeKey(student.name) === oldSelectedName);
        return nextIndex >= 0 ? nextIndex : null;
      })();

      return {
        name: className,
        classSupportNote: String(oldClass?.classSupportNote ?? ""),
        sectionOrder,
        selectedStudentIndex,
        students,
      };
    });

    const activeClassIndex = (() => {
      const oldActiveIndex = Number(existingGrading?.activeClassIndex);
      if (!Number.isInteger(oldActiveIndex) || oldActiveIndex < 0 || oldActiveIndex >= oldClasses.length) {
        return classes.length > 0 ? 0 : null;
      }

      const oldActiveName = normalizeKey(oldClasses[oldActiveIndex]?.name);
      const mappedIndex = classes.findIndex((classItem) => normalizeKey(classItem.name) === oldActiveName);
      if (mappedIndex >= 0) return mappedIndex;
      if (oldActiveIndex < classes.length) return oldActiveIndex;
      return classes.length > 0 ? 0 : null;
    })();

    return {
      sections,
      includeComments: state.setup.includeComments,
      helpThreshold: state.setup.helpThreshold,
      overallProgressNote: String(existingGrading?.overallProgressNote ?? ""),
      activeClassIndex,
      classes,
    };
  };

  const sectionScore = (sectionConfig, studentSection) => {
    if (Number.isFinite(studentSection.overrideScore)) {
      return clamp(studentSection.overrideScore, 0, 100);
    }

    const adjustedScores = [];

    for (let slotIndex = 0; slotIndex < sectionConfig.slots; slotIndex += 1) {
      const score = studentSection.scores[slotIndex];
      if (!Number.isFinite(score)) continue;

      const deduction = sectionConfig.allowDeductions
        ? clamp(numberOr(studentSection.deductions[slotIndex], 0), 0, 100)
        : 0;

      adjustedScores.push(clamp(score - deduction, 0, 100));
    }

    if (adjustedScores.length === 0) return null;

    if (sectionConfig.scoringMode === "highest") {
      return Math.max(...adjustedScores);
    }

    const sum = adjustedScores.reduce((total, value) => total + value, 0);
    return sum / adjustedScores.length;
  };

  const studentFinalScore = (studentRecord) => {
    if (Number.isFinite(studentRecord.totalOverride)) {
      return clamp(studentRecord.totalOverride, 0, 100);
    }

    const totals = state.grading.sections.reduce(
      (accumulator, sectionConfig, sectionIndex) => {
      const score = sectionScore(sectionConfig, studentRecord.sections[sectionIndex]);
      if (score === null) return accumulator;
      accumulator.weightedScore += score * (sectionConfig.weight / 100);
      accumulator.weightUsed += sectionConfig.weight;
      return accumulator;
    },
      { weightedScore: 0, weightUsed: 0 }
    );

    if (totals.weightUsed <= 0) return null;
    return clamp((totals.weightedScore / totals.weightUsed) * 100, 0, 100);
  };

  const studentHasAnyScore = (studentRecord) =>
    Number.isFinite(studentRecord.totalOverride) ||
    studentRecord.sections.some(
      (sectionRecord) =>
        Number.isFinite(sectionRecord.overrideScore) ||
        sectionRecord.scores.some((score) => Number.isFinite(score))
    );

  const normalizedSectionOrder = (ownerRecord, sectionCount) => {
    const rawOrder = Array.isArray(ownerRecord.sectionOrder) ? ownerRecord.sectionOrder : [];
    const uniqueValid = [];
    rawOrder.forEach((value) => {
      const index = Number(value);
      if (!Number.isInteger(index)) return;
      if (index < 0 || index >= sectionCount) return;
      if (uniqueValid.includes(index)) return;
      uniqueValid.push(index);
    });
    for (let index = 0; index < sectionCount; index += 1) {
      if (!uniqueValid.includes(index)) uniqueValid.push(index);
    }
    ownerRecord.sectionOrder = uniqueValid;
    return uniqueValid;
  };

  const moveClassSection = (classRecord, fromSectionIndex, toSectionIndex) => {
    const order = normalizedSectionOrder(classRecord, state.grading.sections.length);
    const fromPosition = order.indexOf(fromSectionIndex);
    const toPosition = order.indexOf(toSectionIndex);
    if (fromPosition < 0 || toPosition < 0 || fromPosition === toPosition) return false;

    const [movedIndex] = order.splice(fromPosition, 1);
    const insertPosition = toPosition;
    order.splice(insertPosition, 0, movedIndex);
    classRecord.sectionOrder = order;
    return true;
  };

  const scoreSelectHtml = (selectedValue) =>
    [
      `<option value="" ${selectedValue === null ? "selected" : ""}>-</option>`,
      ...SCORE_OPTIONS.map((score) => {
        const selected = selectedValue === score ? "selected" : "";
        return `<option value="${score}" ${selected}>${score}</option>`;
      }),
    ].join("");

  const renderSectionGrader = (sectionConfig, sectionRecord, classIndex, studentIndex, sectionIndex) => {
    const slotLimit = 4;
    const expandedKey = `${classIndex}:${studentIndex}:${sectionIndex}`;
    const isExpanded = Boolean(state.ui.expandedSlots[expandedKey]);
    const visibleSlotCount = isExpanded ? sectionConfig.slots : Math.min(sectionConfig.slots, slotLimit);

    const slotsHtml = Array.from({ length: visibleSlotCount }, (_, slotIndex) => {
      const scoreValue = sectionRecord.scores[slotIndex];
      const deductionValue = clamp(Math.round(numberOr(sectionRecord.deductions[slotIndex], 0)), 0, 100);
      const itemLabel = sectionConfig.itemNames[slotIndex] || `Item ${slotIndex + 1}`;
      const slotRowClass = sectionConfig.allowDeductions ? "slot-row" : "slot-row single-score";

      return `
        <div class="${slotRowClass}">
          <div class="slot-title">
            <button
              type="button"
              class="editable-inline"
              data-type="rename-item"
              data-section-index="${sectionIndex}"
              data-slot-index="${slotIndex}"
            >${htmlEscape(itemLabel)}</button>
          </div>
          <label>
            Score
            <select
              data-type="score"
              data-class-index="${classIndex}"
              data-student-index="${studentIndex}"
              data-section-index="${sectionIndex}"
              data-slot-index="${slotIndex}"
            >
              ${scoreSelectHtml(scoreValue)}
            </select>
          </label>
          ${
            sectionConfig.allowDeductions
              ? `
            <label>
              Deduction
              <input
                type="number"
                min="0"
                max="100"
                step="1"
                value="${deductionValue}"
                data-type="deduction"
                data-class-index="${classIndex}"
                data-student-index="${studentIndex}"
                data-section-index="${sectionIndex}"
                data-slot-index="${slotIndex}"
              />
            </label>
          `
              : ""
          }
        </div>
      `;
    }).join("");

    const slotToggleHtml =
      sectionConfig.slots > slotLimit
        ? `
      <button
        type="button"
        class="see-more-btn slot-see-more"
        data-type="toggle-slot-list"
        data-class-index="${classIndex}"
        data-student-index="${studentIndex}"
        data-section-index="${sectionIndex}"
      >${isExpanded ? "Show less" : `See more (${sectionConfig.slots - slotLimit} more)`}</button>
    `
        : "";

    const commentHtml = state.grading.includeComments
      ? `
        <label class="section-comment">
          Comment (optional)
          <textarea
            rows="1"
            data-type="section-comment"
            data-class-index="${classIndex}"
            data-student-index="${studentIndex}"
            data-section-index="${sectionIndex}"
          >${htmlEscape(sectionRecord.comment)}</textarea>
        </label>
      `
      : "";

    return `
      <section
        class="section-box"
        draggable="true"
        data-type="section-card"
        data-class-index="${classIndex}"
        data-student-index="${studentIndex}"
        data-section-index="${sectionIndex}"
      >
        <div class="section-header">
          <h4>
            <button
              type="button"
              class="editable-inline"
              data-type="rename-section"
              data-section-index="${sectionIndex}"
            >${htmlEscape(sectionConfig.name)}</button>
            (${sectionConfig.weight}%)
          </h4>
          <div class="section-head-right">
            <div class="section-item-tools" aria-label="Section item controls">
              <button
                type="button"
                class="mini-icon-btn"
                data-type="add-slot"
                data-section-index="${sectionIndex}"
                title="Add item"
                aria-label="Add item"
              >+</button>
              <button
                type="button"
                class="mini-icon-btn"
                data-type="remove-slot"
                data-section-index="${sectionIndex}"
                title="Remove last item"
                aria-label="Remove last item"
              >-</button>
            </div>
            <div class="section-score">
            Section score:
            <strong
              data-role="section-score"
              data-class-index="${classIndex}"
              data-student-index="${studentIndex}"
              data-section-index="${sectionIndex}"
              title="Double-click to set or clear a manual section score"
            >--</strong>
            </div>
          </div>
        </div>
        <div class="section-body ${state.grading.includeComments ? "" : "no-comment"}">
          <div class="slots-list">
            ${slotsHtml}
            ${slotToggleHtml}
          </div>
          ${commentHtml}
        </div>
      </section>
    `;
  };

  const renderStudentGrader = (studentRecord, classRecord, classIndex, studentIndex) => {
    const sectionOrder = normalizedSectionOrder(classRecord, state.grading.sections.length);
    const sectionsHtml = sectionOrder
      .map((sectionIndex) => {
        const sectionConfig = state.grading.sections[sectionIndex];
        return renderSectionGrader(
          sectionConfig,
          studentRecord.sections[sectionIndex],
          classIndex,
          studentIndex,
          sectionIndex
        );
      })
      .join("");

    return `
      <article class="student-card">
        <div class="student-header">
          <button
            type="button"
            class="editable-inline student-name"
            data-type="rename-student"
            data-class-index="${classIndex}"
            data-student-index="${studentIndex}"
          >${htmlEscape(studentRecord.name)}</button>
          <div class="student-meta">
            <span class="score-chip">
              Final grade:
              <strong
                data-role="final-score"
                data-class-index="${classIndex}"
                data-student-index="${studentIndex}"
                title="Double-click to set or clear a manual total grade"
              >
                --
              </strong>
            </span>
          </div>
        </div>
        <div class="section-strip" data-type="section-strip" data-class-index="${classIndex}" data-student-index="${studentIndex}">
          ${sectionsHtml}
        </div>
      </article>
    `;
  };

  const getFilteredStudentRows = (classRecord) => {
    return classRecord.students
      .map((student, index) => {
        const hasData = studentHasAnyScore(student);
        const scoreText = hasData ? `${studentFinalScore(student).toFixed(1)}%` : "--";
        return {
          index,
          name: student.name,
          label: `${student.name} (${scoreText})`,
        };
      });
  };

  const renderClassGrader = (classRecord, classIndex) => {
    const studentRows = getFilteredStudentRows(classRecord);
    const optionsHtml = studentRows
      .map((row) => {
        const isSelected = row.index === classRecord.selectedStudentIndex ? "selected" : "";
        return `<option value="${row.index}" ${isSelected}>${htmlEscape(row.label)}</option>`;
      })
      .join("");

    const selectedStudent = Number.isInteger(classRecord.selectedStudentIndex)
      ? classRecord.students[classRecord.selectedStudentIndex]
      : null;

    return `
      <section class="class-card" data-class-index="${classIndex}">
        <div class="class-head">
          <div class="class-head-main">
            <h3>
              <button
                type="button"
                class="editable-inline"
                data-type="rename-class"
                data-class-index="${classIndex}"
              >${htmlEscape(classRecord.name)}</button>
            </h3>
            <div class="class-head-actions">
              <button type="button" data-type="add-student" data-class-index="${classIndex}">Add Student</button>
              <button
                type="button"
                class="danger-btn"
                data-type="remove-class"
                data-class-index="${classIndex}"
              >Remove Class</button>
            </div>
          </div>
          <div class="class-metrics">
            <span class="class-stat">Students: ${classRecord.students.length}</span>
            <span class="class-stat">Graded: <strong data-role="class-graded-count" data-class-index="${classIndex}">0</strong></span>
            <span class="class-stat">Average: <strong data-role="class-average" data-class-index="${classIndex}">--</strong></span>
            <span class="class-stat">Need help: <strong data-role="class-help-count" data-class-index="${classIndex}">0</strong></span>
          </div>
        </div>
        <div class="class-bars">
          <div class="class-bar-item">
            <div class="class-bar-top">
              <span>Average progress</span>
              <strong data-role="class-average-label" data-class-index="${classIndex}">--</strong>
            </div>
            <div class="class-bar-track">
              <div class="class-bar-fill" data-role="class-average-bar" data-class-index="${classIndex}"></div>
            </div>
          </div>
          <div class="class-bar-item">
            <div class="class-bar-top">
              <span>Students below threshold</span>
              <strong data-role="class-help-label" data-class-index="${classIndex}">0%</strong>
            </div>
            <div class="class-bar-track">
              <div class="class-bar-fill warn" data-role="class-help-bar" data-class-index="${classIndex}"></div>
            </div>
          </div>
        </div>
        <p class="help-list" data-role="class-help-list" data-class-index="${classIndex}">
          Students needing help: none
        </p>

        <div class="student-pick-grid">
          <label>
            Select student
            <select data-type="student-select" data-class-index="${classIndex}">
              <option value="">Choose a student</option>
              ${optionsHtml}
            </select>
          </label>
          <div class="student-pick-actions">
            <button class="close-sheet-btn" type="button" data-type="clear-student" data-class-index="${classIndex}">Close Sheet</button>
            <button type="button" data-type="rename-selected-student" data-class-index="${classIndex}">Rename Student</button>
            <button type="button" class="danger-btn" data-type="remove-selected-student" data-class-index="${classIndex}">Remove Student</button>
          </div>
        </div>

        ${
          state.grading.includeComments
            ? `
          <div class="grid">
            <label>
              Class support note
              <textarea
                rows="2"
                data-type="class-support-note"
                data-class-index="${classIndex}"
              >${htmlEscape(classRecord.classSupportNote)}</textarea>
            </label>
          </div>
        `
            : ""
        }

        ${
          selectedStudent
            ? renderStudentGrader(
                selectedStudent,
                classRecord,
                classIndex,
                classRecord.selectedStudentIndex
              )
            : '<div class="student-panel-empty">Select a student to open the full grading sheet.</div>'
        }
      </section>
    `;
  };

  const renderSummaryTop = () => {
    const sectionChips = state.grading.sections
      .map((section) => {
        const mode = section.scoringMode === "highest" ? "highest score" : "average";
        const deductionLabel = section.allowDeductions ? "deductions on" : "deductions off";
        return `<span class="section-chip">${htmlEscape(section.name)} | ${section.weight}% | ${section.slots} slots | ${mode} | ${deductionLabel}</span>`;
      })
      .join("");

    summaryTopEl.innerHTML = `
      <div class="summary-top">
        <p class="hint">
          Help threshold: <strong>${state.grading.helpThreshold}%</strong>.
          Blank score slots are ignored inside each section score. Click any student/section/item name to rename it.
        </p>
        <div class="section-chip-row">${sectionChips}</div>
      </div>
    `;
  };

  const renderGrader = () => {
    renderSummaryTop();
    const activeClassIndex = Number(state.grading.activeClassIndex);
    if (
      Number.isInteger(activeClassIndex) &&
      activeClassIndex >= 0 &&
      activeClassIndex < state.grading.classes.length
    ) {
      classesGradingEl.innerHTML = renderClassGrader(
        state.grading.classes[activeClassIndex],
        activeClassIndex
      );
    } else {
      classesGradingEl.innerHTML =
        '<div class="student-panel-empty">Select a class card above to open its grading panel.</div>';
    }
    overallProgressNoteEl.value = state.grading.overallProgressNote;
    if (overallNotesGridEl) {
      overallNotesGridEl.classList.toggle("hidden", !state.grading.includeComments);
    }
    refreshAllMetrics();
  };

  const sectionScoreElement = (classIndex, studentIndex, sectionIndex) =>
    document.querySelector(
      `[data-role="section-score"][data-class-index="${classIndex}"][data-student-index="${studentIndex}"][data-section-index="${sectionIndex}"]`
    );

  const studentScoreElement = (classIndex, studentIndex) =>
    document.querySelector(
      `[data-role="final-score"][data-class-index="${classIndex}"][data-student-index="${studentIndex}"]`
    );

  const classAverageElement = (classIndex) =>
    document.querySelector(`[data-role="class-average"][data-class-index="${classIndex}"]`);

  const classHelpCountElement = (classIndex) =>
    document.querySelector(`[data-role="class-help-count"][data-class-index="${classIndex}"]`);

  const classHelpListElement = (classIndex) =>
    document.querySelector(`[data-role="class-help-list"][data-class-index="${classIndex}"]`);

  const classGradedCountElement = (classIndex) =>
    document.querySelector(`[data-role="class-graded-count"][data-class-index="${classIndex}"]`);

  const classAverageBarElement = (classIndex) =>
    document.querySelector(`[data-role="class-average-bar"][data-class-index="${classIndex}"]`);

  const classAverageLabelElement = (classIndex) =>
    document.querySelector(`[data-role="class-average-label"][data-class-index="${classIndex}"]`);

  const classHelpBarElement = (classIndex) =>
    document.querySelector(`[data-role="class-help-bar"][data-class-index="${classIndex}"]`);

  const classHelpLabelElement = (classIndex) =>
    document.querySelector(`[data-role="class-help-label"][data-class-index="${classIndex}"]`);

  const classStudentSelectElement = (classIndex) =>
    document.querySelector(`[data-type="student-select"][data-class-index="${classIndex}"]`);

  const refreshClassSelector = (classIndex) => {
    const classRecord = state.grading.classes[classIndex];
    if (!classRecord) return;

    const selectEl = classStudentSelectElement(classIndex);
    if (!selectEl) return;

    const rows = getFilteredStudentRows(classRecord);
    const optionsHtml = rows
      .map((row) => `<option value="${row.index}">${htmlEscape(row.label)}</option>`)
      .join("");

    selectEl.innerHTML = `<option value="">Choose a student</option>${optionsHtml}`;
    selectEl.value =
      Number.isInteger(classRecord.selectedStudentIndex) && classRecord.selectedStudentIndex >= 0
        ? String(classRecord.selectedStudentIndex)
        : "";
  };

  const refreshVisibleStudentCard = (classIndex, studentIndex) => {
    const classRecord = state.grading.classes[classIndex];
    const studentRecord = classRecord.students[studentIndex];

    state.grading.sections.forEach((sectionConfig, sectionIndex) => {
      const score = sectionScore(sectionConfig, studentRecord.sections[sectionIndex]);
      const scoreEl = sectionScoreElement(classIndex, studentIndex, sectionIndex);
      if (scoreEl) {
        scoreEl.textContent = score === null ? "--" : `${score.toFixed(1)}%`;
        const isManual = Number.isFinite(studentRecord.sections[sectionIndex].overrideScore);
        scoreEl.classList.toggle("manual", isManual);
        scoreEl.title = isManual
          ? "Manual section score is active. Double-click to edit or clear."
          : "Double-click to set a manual section score.";
      }
    });

    const finalScore = studentFinalScore(studentRecord);
    const finalEl = studentScoreElement(classIndex, studentIndex);
    if (finalEl) {
      finalEl.textContent = finalScore === null ? "--" : `${finalScore.toFixed(1)}%`;
      const isManual = Number.isFinite(studentRecord.totalOverride);
      finalEl.classList.toggle("manual", isManual);
      finalEl.title = isManual
        ? "Manual total grade is active. Double-click to edit or clear."
        : "Double-click to set a manual total grade.";
    }
  };

  const computeClassMetricsData = (classRecord) => {
    const rows = classRecord.students.map((student, studentIndex) => {
      const score = studentFinalScore(student);
      const hasData = studentHasAnyScore(student);
      return { name: student.name, score, hasData, student, studentIndex };
    });

    const gradedRows = rows.filter((row) => row.hasData);
    const needingHelp = gradedRows.filter((row) => row.score < state.grading.helpThreshold);
    const classAverage =
      gradedRows.length === 0
        ? null
        : gradedRows.reduce((sum, row) => sum + row.score, 0) / gradedRows.length;
    const helpRatio = classRecord.students.length === 0 ? 0 : needingHelp.length / classRecord.students.length;

    const sectionAverages = state.grading.sections.map((sectionConfig, sectionIndex) => {
      const sectionRows = classRecord.students
        .map((student) => sectionScore(sectionConfig, student.sections[sectionIndex]))
        .filter((value) => Number.isFinite(value));
      const average =
        sectionRows.length === 0
          ? null
          : sectionRows.reduce((sum, value) => sum + value, 0) / sectionRows.length;
      return {
        name: sectionConfig.name,
        weight: sectionConfig.weight,
        average,
      };
    });

    const sectionAverageValues = sectionAverages
      .map((sectionItem) => sectionItem.average)
      .filter((value) => Number.isFinite(value));
    const averageOfSectionAverages =
      sectionAverageValues.length === 0
        ? null
        : sectionAverageValues.reduce((sum, value) => sum + value, 0) / sectionAverageValues.length;

    return {
      rows,
      gradedRows,
      needingHelp,
      classAverage,
      classSize: classRecord.students.length,
      gradedCount: gradedRows.length,
      helpRatio,
      sectionAverages,
      averageOfSectionAverages,
    };
  };

  const refreshClassMetrics = (classIndex) => {
    const classRecord = state.grading.classes[classIndex];
    const metrics = computeClassMetricsData(classRecord);
    const { gradedRows, needingHelp, classAverage, helpRatio } = metrics;

    if (Number.isInteger(classRecord.selectedStudentIndex)) {
      refreshVisibleStudentCard(classIndex, classRecord.selectedStudentIndex);
    }

    const averageEl = classAverageElement(classIndex);
    if (averageEl) averageEl.textContent = classAverage === null ? "--" : `${classAverage.toFixed(1)}%`;

    const gradedCountEl = classGradedCountElement(classIndex);
    if (gradedCountEl) gradedCountEl.textContent = String(gradedRows.length);

    const helpCountEl = classHelpCountElement(classIndex);
    if (helpCountEl) helpCountEl.textContent = String(needingHelp.length);

    const avgBarEl = classAverageBarElement(classIndex);
    if (avgBarEl) avgBarEl.style.width = `${classAverage === null ? 0 : classAverage}%`;

    const avgLabelEl = classAverageLabelElement(classIndex);
    if (avgLabelEl) avgLabelEl.textContent = classAverage === null ? "--" : `${classAverage.toFixed(1)}%`;

    const helpBarEl = classHelpBarElement(classIndex);
    if (helpBarEl) helpBarEl.style.width = `${Math.min(100, Math.max(0, helpRatio * 100))}%`;

    const helpLabelEl = classHelpLabelElement(classIndex);
    if (helpLabelEl) helpLabelEl.textContent = `${(helpRatio * 100).toFixed(1)}%`;

    const helpListEl = classHelpListElement(classIndex);
    if (helpListEl) {
      const names = needingHelp.map((row) => row.name).join(", ");
      helpListEl.textContent = `Students needing help: ${names || "none"}`;
    }

    return metrics;
  };

  const renderClassSnapshots = (classResults) => {
    if (!classSnapshotGridEl) return;

    const cardsHtml = classResults
      .map((result, classIndex) => {
        const classRecord = state.grading.classes[classIndex];
        const averageText = result.classAverage === null ? "--" : `${result.classAverage.toFixed(1)}%`;
        const active = classIndex === state.grading.activeClassIndex ? "active" : "";
        return `
          <article class="snapshot-card ${active}" data-type="select-class" data-class-index="${classIndex}">
            <div class="snapshot-title">${htmlEscape(classRecord.name)}</div>
            <strong class="snapshot-grade">${averageText}</strong>
            <div class="snapshot-meta">Graded ${result.gradedCount}/${result.classSize}</div>
            <div class="snapshot-meta">Need help ${result.needingHelp.length} (${(
              result.helpRatio * 100
            ).toFixed(1)}%)</div>
            <div class="snapshot-meta">Click to open class details</div>
          </article>
        `;
      })
      .join("");

    classSnapshotGridEl.innerHTML = `
      <div class="class-snapshot-actions">
        <button type="button" data-type="add-class">Add Class</button>
      </div>
      <div class="class-snapshot-list">
        ${cardsHtml}
      </div>
    `;
  };

  const exportStudentMode = () =>
  sheetExportModeSelectEl?.value === "class" ? "class" : "single";

const refreshExportSelectors = () => {
  if (!sheetExportClassSelectEl || !sheetExportStudentSelectEl || !exportStudentSheetBtn || !exportClassReportBtn) {
    return;
  }

  const mode = exportStudentMode();

  if (!state.grading || state.grading.classes.length === 0) {
    sheetExportClassSelectEl.innerHTML = "<option value=\"\">No classes</option>";
    sheetExportStudentSelectEl.innerHTML = "<option value=\"\">No students</option>";
    sheetExportClassSelectEl.disabled = true;
    sheetExportStudentSelectEl.disabled = true;
    exportStudentSheetBtn.disabled = true;
    exportClassReportBtn.disabled = true;
    exportStudentSheetBtn.textContent = "Export Student Sheet";
    return;
  }

  sheetExportClassSelectEl.disabled = false;
  exportClassReportBtn.disabled = false;

  const currentClassIndex = Number(sheetExportClassSelectEl.value);
  const safeClassIndex =
    Number.isInteger(currentClassIndex) &&
    currentClassIndex >= 0 &&
    currentClassIndex < state.grading.classes.length
      ? currentClassIndex
      : Number.isInteger(state.grading.activeClassIndex) &&
          state.grading.activeClassIndex >= 0 &&
          state.grading.activeClassIndex < state.grading.classes.length
        ? state.grading.activeClassIndex
        : 0;

  sheetExportClassSelectEl.innerHTML = state.grading.classes
    .map(
      (classRecord, classIndex) =>
        `<option value="${classIndex}">${htmlEscape(classRecord.name)}</option>`
    )
    .join("");
  sheetExportClassSelectEl.value = String(safeClassIndex);

  const selectedClass = state.grading.classes[safeClassIndex];
  if (!selectedClass || selectedClass.students.length === 0) {
    sheetExportStudentSelectEl.innerHTML = "<option value=\"\">No students</option>";
    sheetExportStudentSelectEl.disabled = true;
    exportStudentSheetBtn.disabled = true;
    exportStudentSheetBtn.textContent = mode === "class" ? "Export Student Sheets" : "Export Student Sheet";
    return;
  }

  const currentStudentIndex = Number(sheetExportStudentSelectEl.value);
  const safeStudentIndex =
    Number.isInteger(currentStudentIndex) &&
    currentStudentIndex >= 0 &&
    currentStudentIndex < selectedClass.students.length
      ? currentStudentIndex
      : Number.isInteger(selectedClass.selectedStudentIndex)
      ? selectedClass.selectedStudentIndex
      : 0;

  sheetExportStudentSelectEl.innerHTML = selectedClass.students
    .map((student, studentIndex) => {
      const suffix = studentHasAnyScore(student) ? ` (${studentFinalScore(student).toFixed(1)}%)` : "";
      return `<option value="${studentIndex}">${htmlEscape(student.name + suffix)}</option>`;
    })
    .join("");

  if (mode === "class") {
    sheetExportStudentSelectEl.disabled = true;
    exportStudentSheetBtn.textContent = "Export Student Sheets";
  } else {
    sheetExportStudentSelectEl.disabled = false;
    exportStudentSheetBtn.textContent = "Export Student Sheet";
  }

  exportStudentSheetBtn.disabled = false;
  sheetExportStudentSelectEl.value = String(clamp(safeStudentIndex, 0, selectedClass.students.length - 1));
};

const selectedExportIndexes = () => {

    const classIndex = Number(sheetExportClassSelectEl?.value);
    const studentIndex = Number(sheetExportStudentSelectEl?.value);
    if (!state.grading) return { classIndex: null, studentIndex: null };
    if (!Number.isInteger(classIndex) || classIndex < 0 || classIndex >= state.grading.classes.length) {
      return { classIndex: null, studentIndex: null };
    }
    const classRecord = state.grading.classes[classIndex];
    const safeStudentIndex =
      Number.isInteger(studentIndex) && studentIndex >= 0 && studentIndex < classRecord.students.length
        ? studentIndex
        : null;
    return { classIndex, studentIndex: safeStudentIndex };
  };

  const isTouchLikeDevice = () =>
    window.matchMedia("(max-width: 900px)").matches ||
    /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);

  const openExportInTab = (url) => {
    const tab = window.open(url, "_blank", "noopener");
    if (!tab) return false;
    setSaveStatus("Opened in a new tab for Share or Print.", "good");
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return true;
  };

  const tryShareHtmlExport = async (filename, blob, successPrefix) => {
    if (typeof navigator.share !== "function") return false;
    try {
      const file = new File([blob], filename, { type: "text/html" });
      if (typeof navigator.canShare === "function" && !navigator.canShare({ files: [file] })) {
        return false;
      }
      await navigator.share({
        title: filename,
        files: [file],
      });
      setSaveStatus(`${successPrefix} Shared from your device.`, "good");
      return true;
    } catch (error) {
      if (error && error.name === "AbortError") {
        setSaveStatus("Share canceled.");
        return true;
      }
      return false;
    }
  };

  const deliverHtmlExport = (filename, htmlBody, successPrefix) => {
    const blob = new Blob([htmlBody], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    if (isTouchLikeDevice()) {
      tryShareHtmlExport(filename, blob, successPrefix).then((shared) => {
        if (shared) {
          URL.revokeObjectURL(url);
          return;
        }
        if (openExportInTab(url)) return;
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        setSaveStatus(`${successPrefix} Downloaded as HTML.`, "good");
      });
      return;
    }

    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setSaveStatus(`${successPrefix} Downloaded as HTML.`, "good");
  };

  const exportBaseStyle = `
    <style>
      :root {
        --teal: #167a78;
        --teal-soft: #e5f5f4;
        --bronze: #b1834f;
        --bronze-soft: #f2e6d8;
        --ink: #173331;
        --muted: #4c6663;
        --line: #c4d8d7;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        padding: 26px;
        color: var(--ink);
        font-family: "Avenir Next", "Trebuchet MS", "Verdana", sans-serif;
        background: linear-gradient(180deg, #f5fbfb 0%, #e9f4f4 100%);
      }
      .wrap { max-width: 940px; margin: 0 auto; display: grid; gap: 14px; }
      .head {
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 14px;
        background: #fff;
      }
      .chip-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
      .chip {
        border-radius: 999px;
        border: 1px solid #d2c2ad;
        background: var(--bronze-soft);
        color: #725432;
        padding: 5px 10px;
        font-size: 12px;
      }
      .metric-grid {
        display: grid;
        gap: 10px;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      }
      .metric {
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 12px;
        background: linear-gradient(130deg, #f7fcfc 0%, #eaf5f4 100%);
      }
      .metric span { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: 0.07em; }
      .metric strong { display: block; font-size: 24px; margin-top: 6px; color: #15615f; }
      .card {
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 12px;
        background: #fff;
      }
      table { width: 100%; border-collapse: collapse; font-size: 13px; }
      th, td { border: 1px solid var(--line); padding: 8px; text-align: left; }
      th { background: #edf6f5; }
      .muted { color: var(--muted); }
      .help-pill {
        display: inline-block;
        border-radius: 999px;
        border: 1px solid #e4c7c2;
        background: #fff2ef;
        color: #873526;
        padding: 4px 10px;
        font-size: 12px;
        margin: 0 6px 6px 0;
      }
      .section-grid { display: grid; gap: 8px; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); }
      .section-box {
        border: 1px solid var(--line);
        border-radius: 10px;
        padding: 10px;
        background: #fbfefe;
      }
      .section-box strong { font-size: 20px; color: #165f5d; }
    </style>
  `;

  const exportStudentSheet = (classIndex, studentIndex) => {
    const classRecord = state.grading.classes[classIndex];
    const studentRecord = classRecord.students[studentIndex];
    const classMetrics = computeClassMetricsData(classRecord);

    const sectionRows = state.grading.sections
      .map((sectionConfig, sectionIndex) => {
        const sectionRecord = studentRecord.sections[sectionIndex];
        const score = sectionScore(sectionConfig, sectionRecord);
        const comment = sectionRecord.comment.trim();
        return `
          <tr>
            <td>${htmlEscape(sectionConfig.name)}</td>
            <td>${score === null ? "--" : `${score.toFixed(1)}%`}</td>
            <td>${sectionConfig.weight}%</td>
            <td>${comment ? htmlEscape(comment) : "<span class=\"muted\">No comment</span>"}</td>
          </tr>
        `;
      })
      .join("");

    const finalGrade = studentFinalScore(studentRecord);
    const classAverageText =
      classMetrics.classAverage === null ? "--" : `${classMetrics.classAverage.toFixed(1)}%`;
    const finalGradeText = finalGrade === null ? "--" : `${finalGrade.toFixed(1)}%`;

    const html = `
      <!doctype html>
      <html>
      <head>
        <meta charset="utf-8" />
        <title>${htmlEscape(studentRecord.name)} - Grade Sheet</title>
        ${exportBaseStyle}
      </head>
      <body>
        <main class="wrap">
          <section class="head">
            <h1>${htmlEscape(studentRecord.name)} - Grade Sheet</h1>
            <p class="muted">${htmlEscape(classRecord.name)} | Exported ${htmlEscape(new Date().toLocaleString())}</p>
          </section>
          <section class="metric-grid">
            <article class="metric"><span>Final grade</span><strong>${finalGradeText}</strong></article>
            <article class="metric"><span>Class average</span><strong>${classAverageText}</strong></article>
          </section>
          <section class="card">
            <h2>Section Breakdown and Comments</h2>
            <table>
              <thead>
                <tr>
                  <th>Section</th>
                  <th>Section Score</th>
                  <th>Weight</th>
                  <th>Comment</th>
                </tr>
              </thead>
              <tbody>${sectionRows}</tbody>
            </table>
          </section>
        </main>
      </body>
      </html>
    `;

    const filenameSafeName = studentRecord.name.replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "");
    deliverHtmlExport(`${filenameSafeName || "student"}-grade-sheet.html`, html, `Exported student sheet for ${studentRecord.name}.`);
  };

const exportStudentSheetsByClass = (classIndex) => {
  const classRecord = state.grading.classes[classIndex];
  const classMetrics = computeClassMetricsData(classRecord);
  const classAverageText =
    classMetrics.classAverage === null ? "--" : `${classMetrics.classAverage.toFixed(1)}%`;

  const studentSheetsHtml = classRecord.students
    .map((studentRecord) => {
      const sectionRows = state.grading.sections
        .map((sectionConfig, sectionIndex) => {
          const sectionRecord = studentRecord.sections[sectionIndex];
          const score = sectionScore(sectionConfig, sectionRecord);
          const comment = sectionRecord.comment.trim();
          return `
            <tr>
              <td>${htmlEscape(sectionConfig.name)}</td>
              <td>${score === null ? "--" : `${score.toFixed(1)}%`}</td>
              <td>${sectionConfig.weight}%</td>
              <td>${comment ? htmlEscape(comment) : "<span class=\"muted\">No comment</span>"}</td>
            </tr>
          `;
        })
        .join("");

      const finalGrade = studentFinalScore(studentRecord);
      const finalGradeText = finalGrade === null ? "--" : `${finalGrade.toFixed(1)}%`;

      return `
        <section class="card student-sheet-page">
          <h2>${htmlEscape(studentRecord.name)}</h2>
          <p class="muted">${htmlEscape(classRecord.name)} | Class average ${classAverageText}</p>
          <section class="metric-grid">
            <article class="metric"><span>Final grade</span><strong>${finalGradeText}</strong></article>
            <article class="metric"><span>Class average</span><strong>${classAverageText}</strong></article>
          </section>
          <section class="card">
            <h3>Section Breakdown and Comments</h3>
            <table>
              <thead>
                <tr>
                  <th>Section</th>
                  <th>Section Score</th>
                  <th>Weight</th>
                  <th>Comment</th>
                </tr>
              </thead>
              <tbody>${sectionRows}</tbody>
            </table>
          </section>
        </section>
      `;
    })
    .join("");

  const html = `
    <!doctype html>
    <html>
    <head>
      <meta charset="utf-8" />
      <title>${htmlEscape(classRecord.name)} - Student Sheets</title>
      ${exportBaseStyle}
      <style>
        .student-sheet-page { break-after: page; }
        .student-sheet-page:last-child { break-after: auto; }
      </style>
    </head>
    <body>
      <main class="wrap">
        <section class="head">
          <h1>${htmlEscape(classRecord.name)} - Student Sheets</h1>
          <p class="muted">Exported ${htmlEscape(new Date().toLocaleString())}</p>
        </section>
        ${studentSheetsHtml}
      </main>
    </body>
    </html>
  `;

  const filenameSafeName = classRecord.name.replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "");
  deliverHtmlExport(
    `${filenameSafeName || "class"}-student-sheets.html`,
    html,
    `Exported student sheets for ${classRecord.name}.`
  );
};

const exportClassReport = (classIndex) => {

    const classRecord = state.grading.classes[classIndex];
    const classMetrics = computeClassMetricsData(classRecord);

    const studentRows = classRecord.students
      .map((student) => {
        const hasData = studentHasAnyScore(student);
        const finalScore = hasData ? `${studentFinalScore(student).toFixed(1)}%` : "--";
        return `
          <tr>
            <td>${htmlEscape(student.name)}</td>
            <td>${finalScore}</td>
            <td>${hasData ? "Yes" : "No"}</td>
          </tr>
        `;
      })
      .join("");

    const helpPills =
      classMetrics.needingHelp.length === 0
        ? '<span class="muted">No students below threshold.</span>'
        : classMetrics.needingHelp
            .map(
              (row) =>
                `<span class="help-pill">${htmlEscape(row.name)} (${row.score.toFixed(1)}%)</span>`
            )
            .join("");

    const classAverageText =
      classMetrics.classAverage === null ? "--" : `${classMetrics.classAverage.toFixed(1)}%`;
    const sectionAverageText =
      classMetrics.averageOfSectionAverages === null
        ? "--"
        : `${classMetrics.averageOfSectionAverages.toFixed(1)}%`;

    const html = `
      <!doctype html>
      <html>
      <head>
        <meta charset="utf-8" />
        <title>${htmlEscape(classRecord.name)} - Class Report</title>
        ${exportBaseStyle}
      </head>
      <body>
        <main class="wrap">
          <section class="head">
            <h1>${htmlEscape(classRecord.name)} - Class Report</h1>
            <p class="muted">Exported ${htmlEscape(new Date().toLocaleString())}</p>
          </section>
          <section class="metric-grid">
            <article class="metric"><span>Class average</span><strong>${classAverageText}</strong></article>
            <article class="metric"><span>Average of section averages</span><strong>${sectionAverageText}</strong></article>
            <article class="metric"><span>Students graded</span><strong>${classMetrics.gradedCount}/${classMetrics.classSize}</strong></article>
            <article class="metric"><span>Need help</span><strong>${classMetrics.needingHelp.length}</strong></article>
          </section>
          <section class="card">
            <h2>Students Needing Help</h2>
            <div>${helpPills}</div>
          </section>
          <section class="card">
            <h2>Quick Student Summary</h2>
            <table>
              <thead>
                <tr>
                  <th>Student</th>
                  <th>Final Grade</th>
                  <th>Has Scores</th>
                </tr>
              </thead>
              <tbody>${studentRows}</tbody>
            </table>
          </section>
          <section class="card">
            <h2>Notes</h2>
            <p><strong>Overall progress:</strong> ${htmlEscape(
              state.grading.overallProgressNote || "No note"
            )}</p>
            <p><strong>Class support:</strong> ${htmlEscape(classRecord.classSupportNote || "No note")}</p>
          </section>
        </main>
      </body>
      </html>
    `;

    const filenameSafeName = classRecord.name.replace(/[^a-z0-9]+/gi, "_").replace(/^_+|_+$/g, "");
    deliverHtmlExport(`${filenameSafeName || "class"}-report.html`, html, `Exported class report for ${classRecord.name}.`);
  };

  const refreshOverallMetrics = () => {
    const classResults = state.grading.classes.map((_, classIndex) => refreshClassMetrics(classIndex));
    renderClassSnapshots(classResults);

    const gradedRows = classResults.flatMap((result) => result.gradedRows);
    const needingHelpRows = classResults.flatMap((result, classIndex) =>
      result.needingHelp.map((row) => ({
        classIndex,
        studentIndex: row.studentIndex,
        className: state.grading.classes[classIndex].name,
        studentName: row.name,
        score: row.score,
      }))
    );

    const overallAverage =
      gradedRows.length === 0
        ? null
        : gradedRows.reduce((sum, row) => sum + row.score, 0) / gradedRows.length;

    overallAverageEl.textContent = overallAverage === null ? "--" : `${overallAverage.toFixed(1)}%`;
    overallHelpCountEl.textContent = String(needingHelpRows.length);
    overallGradedCountEl.textContent = String(gradedRows.length);
    overallClassCountEl.textContent = String(state.grading.classes.length);

    if (needingHelpRows.length === 0) {
      state.ui.showAllHelp = false;
      overallHelpListEl.innerHTML = '<span class="help-pill empty">No students below threshold.</span>';
      return;
    }

    if (needingHelpRows.length <= 7) {
      state.ui.showAllHelp = false;
    }

    const visibleRows = state.ui.showAllHelp ? needingHelpRows : needingHelpRows.slice(0, 7);
    const pillsHtml = visibleRows
      .map(
        (item) =>
          `<button
            type="button"
            class="help-pill"
            data-type="open-help-student"
            data-class-index="${item.classIndex}"
            data-student-index="${item.studentIndex}"
          >${htmlEscape(item.studentName)} - ${htmlEscape(item.className)} (${item.score.toFixed(1)}%)</button>`
      )
      .join("");

    const toggleHtml =
      needingHelpRows.length > 7
        ? `<button type="button" class="see-more-btn" data-type="toggle-help-list">${
            state.ui.showAllHelp
              ? "Show less"
              : `See more (${needingHelpRows.length - 7} more)`
          }</button>`
        : "";

    overallHelpListEl.innerHTML = `${pillsHtml}${toggleHtml}`;
  };

  const refreshAllMetrics = () => {
    refreshOverallMetrics();
    refreshExportSelectors();
  };

  const readIndexes = (target) => ({
    classIndex: Number(target.dataset.classIndex),
    studentIndex: Number(target.dataset.studentIndex),
    sectionIndex: Number(target.dataset.sectionIndex),
    slotIndex: Number(target.dataset.slotIndex),
  });

  const updateSectionConfigFromInput = (target) => {
    const sectionIndex = Number(target.dataset.sectionIndex);
    if (!Number.isInteger(sectionIndex)) return;

    const section = state.setup.sections[sectionIndex];
    if (!section) return;

    const field = target.dataset.field;
    if (field === "name") section.name = target.value;
    if (field === "weight") section.weight = clamp(numberOr(target.value, 0), 0, 100);
    if (field === "slots") section.slots = clamp(Math.round(numberOr(target.value, 1)), 1, 12);
    if (field === "scoringMode") section.scoringMode = target.value === "highest" ? "highest" : "average";
    if (field === "allowDeductions") section.allowDeductions = target.checked;

    updateWeightMessage();
  };

  const updateClassConfigFromInput = (target) => {
    const classIndex = Number(target.dataset.classIndex);
    if (!Number.isInteger(classIndex)) return;

    const classItem = state.setup.classes[classIndex];
    if (!classItem) return;

    const field = target.dataset.field;
    if (field === "name") classItem.name = target.value;
    if (field === "studentCount") classItem.studentCount = clamp(Math.round(numberOr(target.value, 1)), 1, 60);
    if (field === "studentNamesText") classItem.studentNamesText = target.value;
  };

  const snapshotForSave = () => ({
    version: 2,
    savedAt: new Date().toISOString(),
    activeView: setupViewEl.classList.contains("hidden") ? "grader" : "setup",
    setup:
      state.grading && setupViewEl.classList.contains("hidden")
        ? buildSetupFromGrading(state.grading)
        : cloneJson(state.setup),
    grading: state.grading ? cloneJson(state.grading) : null,
  });

  const saveToLocal = (prefix = "Saved in this browser", silent = false) => {
    try {
      const snapshot = snapshotForSave();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
      setLastSavedText(snapshot.savedAt);
      if (!silent && prefix) setSaveStatus(prefix, "good");
      return true;
    } catch (error) {
      if (!silent) setSaveStatus("Could not save locally. Use Download Backup instead.", "bad");
      return false;
    }
  };

  const scheduleAutoSave = () => {
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => {
      saveToLocal("", true);
    }, 450);
  };

  const applySnapshot = (snapshot) => {
    if (!snapshot || typeof snapshot !== "object") {
      setSaveStatus("Saved file is not valid.", "bad");
      return false;
    }

    normalizeSetup(snapshot.setup);
    state.grading = normalizeGrading(snapshot.grading);
    state.ui.showAllHelp = false;

    renderSetup();

    if (state.grading) {
      renderGrader();
      if (snapshot.activeView === "setup") showSetupView();
      else showGraderView();
    } else {
      showSetupView();
      refreshExportSelectors();
    }

    const stampText = snapshot.savedAt ? formatStamp(snapshot.savedAt) : "";
    setLastSavedText(snapshot.savedAt);
    setSaveStatus(`Loaded ${stampText ? `(${stampText})` : "saved data"}`, "good");
    return true;
  };

  const loadFromLocal = (silentIfMissing = false) => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (!saved) {
        if (!silentIfMissing) setSaveStatus("No local save found on this browser yet.", "bad");
        return false;
      }

      const snapshot = JSON.parse(saved);
      return applySnapshot(snapshot);
    } catch (error) {
      if (!silentIfMissing) setSaveStatus("Could not load saved data.", "bad");
      return false;
    }
  };

  const exportToFile = () => {
    const snapshot = snapshotForSave();
    const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const stamp = new Date().toISOString().replace(/[:]/g, "-").replace(/\..+$/, "");
    link.href = url;
    link.download = `rubric-grade-save-${stamp}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setLastSavedText(snapshot.savedAt);
    setSaveStatus(`Exported ${formatStamp(snapshot.savedAt)}`, "good");
  };

  const importFromFile = async (file) => {
    if (!file) return;
    try {
      const text = await file.text();
      const snapshot = JSON.parse(text);
      if (!applySnapshot(snapshot)) return;
      if (importRequestedFromSetup && state.grading) {
        showGraderView();
        renderGrader();
      }
      saveToLocal("Imported and saved", false);
    } catch (error) {
      setSaveStatus("Could not import that file.", "bad");
    }
  };

  const onCreateSheet = () => {
    setupErrorEl.textContent = "";
    ensureSetupArrays();

    const errors = validateSetup();
    if (errors.length > 0) {
      setupErrorEl.textContent = errors[0];
      return;
    }

    state.grading = state.grading
      ? mergeExistingGradingWithSetup(state.grading)
      : buildGradingState();
    state.ui.showAllHelp = false;
    setSaveStatus("Setup changes applied. Existing grades were kept where possible.", "good");
    renderGrader();
    showGraderView();
    scheduleAutoSave();
  };

  sectionCountEl.addEventListener("input", () => {
    state.setup.sectionCount = clamp(Math.round(numberOr(sectionCountEl.value, 1)), 1, 12);
    renderSetup();
    scheduleAutoSave();
  });

  classCountEl.addEventListener("input", () => {
    state.setup.classCount = clamp(Math.round(numberOr(classCountEl.value, 1)), 1, 12);
    renderSetup();
    scheduleAutoSave();
  });

  helpThresholdEl.addEventListener("input", () => {
    state.setup.helpThreshold = clamp(Math.round(numberOr(helpThresholdEl.value, 70)), 0, 100);
    scheduleAutoSave();
  });

  includeCommentsEl.addEventListener("change", () => {
    state.setup.includeComments = includeCommentsEl.checked;
    if (state.grading) state.grading.includeComments = state.setup.includeComments;
    scheduleAutoSave();
  });

  sectionsConfigEl.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.dataset.kind !== "section") return;
    updateSectionConfigFromInput(target);
    scheduleAutoSave();
  });

  sectionsConfigEl.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.dataset.kind !== "section") return;
    updateSectionConfigFromInput(target);
    scheduleAutoSave();
  });

  classesConfigEl.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.dataset.kind !== "class") return;
    updateClassConfigFromInput(target);
    scheduleAutoSave();
  });

  classesConfigEl.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.dataset.kind !== "class") return;
    updateClassConfigFromInput(target);
    scheduleAutoSave();
  });

  createSheetBtn.addEventListener("click", onCreateSheet);

  editSetupBtn.addEventListener("click", async () => {
    if (state.grading) {
      const accepted = await confirmSetupWarning();
      if (!accepted) return;
      syncSetupFromGrading();
    }
    showSetupView();
    renderSetup();
    scheduleAutoSave();
  });

  if (backToGraderBtn) {
    backToGraderBtn.addEventListener("click", () => {
      if (!state.grading) return;
      showGraderView();
      renderGrader();
      scheduleAutoSave();
    });
  }

  classesGradingEl.addEventListener("dblclick", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const role = target.dataset.role;

    if (role === "section-score") {
      const classIndex = Number(target.dataset.classIndex);
      const studentIndex = Number(target.dataset.studentIndex);
      const sectionIndex = Number(target.dataset.sectionIndex);

      const sectionRecord = state.grading.classes[classIndex]?.students[studentIndex]?.sections[sectionIndex];
      const sectionConfig = state.grading.sections[sectionIndex];
      if (!sectionRecord || !sectionConfig) return;

      const currentManual = Number.isFinite(sectionRecord.overrideScore)
        ? sectionRecord.overrideScore
        : null;
      const currentAuto = sectionScore(sectionConfig, sectionRecord);
      const defaultValue =
        currentManual !== null ? String(currentManual) : currentAuto === null ? "" : currentAuto.toFixed(1);

      const response = window.prompt(
        "Manual section score (0-100). Leave blank to return to auto-calculation.",
        defaultValue
      );

      if (response === null) return;
      const clean = response.trim();
      sectionRecord.overrideScore = clean === "" ? null : clamp(numberOr(clean, 0), 0, 100);

      refreshOverallMetrics();
      refreshClassSelector(classIndex);
      scheduleAutoSave();
      return;
    }

    if (role === "final-score") {
      const classIndex = Number(target.dataset.classIndex);
      const studentIndex = Number(target.dataset.studentIndex);
      const studentRecord = state.grading.classes[classIndex]?.students[studentIndex];
      if (!studentRecord) return;

      const currentManual = Number.isFinite(studentRecord.totalOverride) ? studentRecord.totalOverride : null;
      const currentAuto = (() => {
        const totals = state.grading.sections.reduce(
          (accumulator, sectionConfig, sectionIndex) => {
            const score = sectionScore(sectionConfig, studentRecord.sections[sectionIndex]);
            if (score === null) return accumulator;
            accumulator.weightedScore += score * (sectionConfig.weight / 100);
            accumulator.weightUsed += sectionConfig.weight;
            return accumulator;
          },
          { weightedScore: 0, weightUsed: 0 }
        );
        if (totals.weightUsed <= 0) return null;
        return (totals.weightedScore / totals.weightUsed) * 100;
      })();
      const defaultValue =
        currentManual !== null ? String(currentManual) : currentAuto === null ? "" : currentAuto.toFixed(1);

      const response = window.prompt(
        "Manual total grade (0-100). Leave blank to return to auto-calculation.",
        defaultValue
      );
      if (response === null) return;
      const clean = response.trim();
      studentRecord.totalOverride = clean === "" ? null : clamp(numberOr(clean, 0), 0, 100);

      refreshOverallMetrics();
      refreshClassSelector(classIndex);
      scheduleAutoSave();
    }
  });

  classesGradingEl.addEventListener("click", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.dataset.type === "rename-class") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      const nextName = window.prompt("Rename class", classRecord.name);
      if (nextName === null) return;
      const cleanName = nextName.trim();
      if (!cleanName) return;
      classRecord.name = cleanName;
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "add-student") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      const studentName = nextStudentName(classRecord);
      classRecord.students.push(createBlankStudentRecord(studentName));
      classRecord.selectedStudentIndex = classRecord.students.length - 1;
      renderGrader();
      setSaveStatus(`Added ${studentName} to ${classRecord.name}`, "good");
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "remove-class") {
      const classIndex = Number(target.dataset.classIndex);
      if (!Number.isInteger(classIndex) || classIndex < 0 || classIndex >= state.grading.classes.length) return;
      if (state.grading.classes.length <= 1) {
        setSaveStatus("At least one class is required. Add another class first.", "bad");
        return;
      }

      const className = state.grading.classes[classIndex].name;
      const confirmed = window.confirm(`Remove class "${className}" and all its student grades?`);
      if (!confirmed) return;

      state.grading.classes.splice(classIndex, 1);
      const active = Number(state.grading.activeClassIndex);
      if (!Number.isInteger(active)) {
        state.grading.activeClassIndex = 0;
      } else if (active > classIndex) {
        state.grading.activeClassIndex = active - 1;
      } else if (active === classIndex) {
        state.grading.activeClassIndex = Math.min(classIndex, state.grading.classes.length - 1);
      }

      renderGrader();
      setSaveStatus(`Removed ${className}`, "good");
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "rename-selected-student") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      const studentIndex = Number(classRecord.selectedStudentIndex);
      if (!Number.isInteger(studentIndex) || studentIndex < 0 || studentIndex >= classRecord.students.length) {
        setSaveStatus("Select a student first.", "bad");
        return;
      }
      const studentRecord = classRecord.students[studentIndex];
      const nextName = window.prompt("Rename student", studentRecord.name);
      if (nextName === null) return;
      const cleanName = nextName.trim();
      if (!cleanName) return;
      studentRecord.name = cleanName;
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "remove-selected-student") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      const studentIndex = Number(classRecord.selectedStudentIndex);
      if (!Number.isInteger(studentIndex) || studentIndex < 0 || studentIndex >= classRecord.students.length) {
        setSaveStatus("Select a student first.", "bad");
        return;
      }
      if (classRecord.students.length <= 1) {
        setSaveStatus("At least one student is required in a class.", "bad");
        return;
      }

      const studentName = classRecord.students[studentIndex].name;
      const confirmed = window.confirm(`Remove student "${studentName}" from ${classRecord.name}?`);
      if (!confirmed) return;

      classRecord.students.splice(studentIndex, 1);
      classRecord.selectedStudentIndex = Math.min(studentIndex, classRecord.students.length - 1);
      renderGrader();
      setSaveStatus(`Removed ${studentName}`, "good");
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "rename-student") {
      const classIndex = Number(target.dataset.classIndex);
      const studentIndex = Number(target.dataset.studentIndex);
      const classRecord = state.grading.classes[classIndex];
      const studentRecord = classRecord?.students[studentIndex];
      if (!studentRecord) return;
      const nextName = window.prompt("Rename student", studentRecord.name);
      if (nextName === null) return;
      const cleanName = nextName.trim();
      if (!cleanName) return;
      studentRecord.name = cleanName;
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "rename-section") {
      const sectionIndex = Number(target.dataset.sectionIndex);
      const sectionConfig = state.grading.sections[sectionIndex];
      if (!sectionConfig) return;
      const nextName = window.prompt("Rename section", sectionConfig.name);
      if (nextName === null) return;
      const cleanName = nextName.trim();
      if (!cleanName) return;
      sectionConfig.name = cleanName;
      if (state.setup.sections[sectionIndex]) state.setup.sections[sectionIndex].name = cleanName;
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "rename-item") {
      const sectionIndex = Number(target.dataset.sectionIndex);
      const slotIndex = Number(target.dataset.slotIndex);
      const sectionConfig = state.grading.sections[sectionIndex];
      if (!sectionConfig) return;
      const currentName = sectionConfig.itemNames[slotIndex] || `Item ${slotIndex + 1}`;
      const nextName = window.prompt("Rename item", currentName);
      if (nextName === null) return;
      const cleanName = nextName.trim();
      if (!cleanName) return;
      sectionConfig.itemNames[slotIndex] = cleanName;
      if (state.setup.sections[sectionIndex]) {
        if (!Array.isArray(state.setup.sections[sectionIndex].itemNames)) {
          state.setup.sections[sectionIndex].itemNames = Array(sectionConfig.slots).fill("");
        }
        state.setup.sections[sectionIndex].itemNames[slotIndex] = cleanName;
      }
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "add-slot") {
      const sectionIndex = Number(target.dataset.sectionIndex);
      if (!Number.isInteger(sectionIndex)) return;
      const added = addSectionItemSlot(sectionIndex);
      if (!added) {
        setSaveStatus("Each section can have up to 12 items.", "bad");
        return;
      }
      renderGrader();
      setSaveStatus("Added section item.", "good");
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "remove-slot") {
      const sectionIndex = Number(target.dataset.sectionIndex);
      if (!Number.isInteger(sectionIndex)) return;
      const removed = removeSectionItemSlot(sectionIndex);
      if (!removed.ok) {
        if (removed.reason === "min") setSaveStatus("Each section needs at least 1 item.", "bad");
        return;
      }
      renderGrader();
      setSaveStatus("Removed last section item.", "good");
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type === "toggle-slot-list") {
      const classIndex = Number(target.dataset.classIndex);
      const studentIndex = Number(target.dataset.studentIndex);
      const sectionIndex = Number(target.dataset.sectionIndex);
      if (![classIndex, studentIndex, sectionIndex].every(Number.isInteger)) return;
      const key = `${classIndex}:${studentIndex}:${sectionIndex}`;
      state.ui.expandedSlots[key] = !state.ui.expandedSlots[key];
      renderGrader();
      return;
    }

    if (target.dataset.type === "clear-student") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      classRecord.selectedStudentIndex = null;
      renderGrader();
      scheduleAutoSave();
    }
  });

  classesGradingEl.addEventListener("dragstart", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const cardEl = target.closest("[data-type='section-card']");
    if (!(cardEl instanceof HTMLElement)) return;

    const classIndex = Number(cardEl.dataset.classIndex);
    const studentIndex = Number(cardEl.dataset.studentIndex);
    const sectionIndex = Number(cardEl.dataset.sectionIndex);
    if (![classIndex, studentIndex, sectionIndex].every(Number.isInteger)) return;

    state.ui.dragSource = { classIndex, studentIndex, sectionIndex };
    cardEl.classList.add("dragging");

    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", String(sectionIndex));
    }
  });

  classesGradingEl.addEventListener("dragover", (event) => {
    if (!state.grading || !state.ui.dragSource) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const cardEl = target.closest("[data-type='section-card']");
    if (!(cardEl instanceof HTMLElement)) return;

    const classIndex = Number(cardEl.dataset.classIndex);
    const studentIndex = Number(cardEl.dataset.studentIndex);
    if (
      classIndex !== state.ui.dragSource.classIndex ||
      studentIndex !== state.ui.dragSource.studentIndex
    ) {
      return;
    }

    event.preventDefault();
    document
      .querySelectorAll("[data-type='section-card'].drop-target")
      .forEach((node) => node.classList.remove("drop-target"));
    cardEl.classList.add("drop-target");
    if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  });

  classesGradingEl.addEventListener("dragleave", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const cardEl = target.closest("[data-type='section-card']");
    if (!(cardEl instanceof HTMLElement)) return;
    cardEl.classList.remove("drop-target");
  });

  classesGradingEl.addEventListener("drop", (event) => {
    if (!state.grading || !state.ui.dragSource) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const cardEl = target.closest("[data-type='section-card']");
    if (!(cardEl instanceof HTMLElement)) return;
    event.preventDefault();

    const classIndex = Number(cardEl.dataset.classIndex);
    const studentIndex = Number(cardEl.dataset.studentIndex);
    const sectionIndex = Number(cardEl.dataset.sectionIndex);
    if (![classIndex, studentIndex, sectionIndex].every(Number.isInteger)) return;

    const dragSource = state.ui.dragSource;
    if (
      classIndex !== dragSource.classIndex ||
      studentIndex !== dragSource.studentIndex ||
      sectionIndex === dragSource.sectionIndex
    ) {
      cardEl.classList.remove("drop-target");
      return;
    }

    const classRecord = state.grading.classes[classIndex];
    if (!classRecord) return;
    const changed = moveClassSection(classRecord, dragSource.sectionIndex, sectionIndex);
    cardEl.classList.remove("drop-target");
    if (!changed) return;

    state.ui.dragSource = null;
    renderGrader();
    scheduleAutoSave();
  });

  classesGradingEl.addEventListener("dragend", () => {
    state.ui.dragSource = null;
    document.querySelectorAll("[data-type='section-card'].dragging,[data-type='section-card'].drop-target").forEach((node) => {
      node.classList.remove("dragging", "drop-target");
    });
  });

  classesGradingEl.addEventListener("change", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const inputType = target.dataset.type;
    if (!inputType) return;

    if (inputType === "student-select") {
      const classIndex = Number(target.dataset.classIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      classRecord.selectedStudentIndex = target.value === "" ? null : Number(target.value);
      if (Number.isInteger(classRecord.selectedStudentIndex)) {
        collapseOverallSummary();
      }
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (inputType === "score") {
      const { classIndex, studentIndex, sectionIndex, slotIndex } = readIndexes(target);
      const scoreValue = target.value === "" ? null : clamp(numberOr(target.value, 0), 0, 100);
      state.grading.classes[classIndex].students[studentIndex].sections[sectionIndex].scores[slotIndex] = scoreValue;
      refreshOverallMetrics();
      refreshClassSelector(classIndex);
      scheduleAutoSave();
      return;
    }

    if (inputType === "deduction") {
      const { classIndex, studentIndex, sectionIndex, slotIndex } = readIndexes(target);
      const deductionValue = clamp(Math.round(numberOr(target.value, 0)), 0, 100);
      state.grading.classes[classIndex].students[studentIndex].sections[sectionIndex].deductions[slotIndex] =
        deductionValue;
      target.value = String(deductionValue);
      refreshOverallMetrics();
      refreshClassSelector(classIndex);
      scheduleAutoSave();
    }
  });

  classesGradingEl.addEventListener("keydown", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) return;
    if (target.dataset.type !== "student-select") return;
    if (event.altKey || event.ctrlKey || event.metaKey) return;

    const now = Date.now();
    const classIndex = Number(target.dataset.classIndex);
    const key = event.key;

    if (key === "Escape") {
      state.ui.studentSelectTypeahead = { classIndex: null, buffer: "", stamp: 0 };
      return;
    }

    if (key === "Backspace") {
      if (now - state.ui.studentSelectTypeahead.stamp > 850) {
        state.ui.studentSelectTypeahead.buffer = "";
      }
      state.ui.studentSelectTypeahead.classIndex = classIndex;
      state.ui.studentSelectTypeahead.buffer = state.ui.studentSelectTypeahead.buffer.slice(0, -1);
      state.ui.studentSelectTypeahead.stamp = now;
      return;
    }

    if (key.length !== 1) return;

    const sameClass = state.ui.studentSelectTypeahead.classIndex === classIndex;
    const recent = now - state.ui.studentSelectTypeahead.stamp <= 850;
    const base = sameClass && recent ? state.ui.studentSelectTypeahead.buffer : "";
    const search = `${base}${key.toLowerCase()}`;

    state.ui.studentSelectTypeahead = {
      classIndex,
      buffer: search,
      stamp: now,
    };

    const match = Array.from(target.options).find((option, optionIndex) => {
      if (optionIndex === 0) return false;
      return option.textContent.toLowerCase().includes(search);
    });
    if (!match) return;

    target.value = match.value;
    target.dispatchEvent(new Event("change", { bubbles: true }));
    event.preventDefault();
  });

  classesGradingEl.addEventListener("input", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const inputType = target.dataset.type;
    if (!inputType) return;

    if (inputType === "deduction") {
      const { classIndex, studentIndex, sectionIndex, slotIndex } = readIndexes(target);
      const deductionValue = clamp(Math.round(numberOr(target.value, 0)), 0, 100);
      state.grading.classes[classIndex].students[studentIndex].sections[sectionIndex].deductions[slotIndex] =
        deductionValue;
      refreshOverallMetrics();
      refreshClassSelector(classIndex);
      scheduleAutoSave();
      return;
    }

    if (inputType === "section-comment") {
      const { classIndex, studentIndex, sectionIndex } = readIndexes(target);
      state.grading.classes[classIndex].students[studentIndex].sections[sectionIndex].comment = target.value;
      scheduleAutoSave();
      return;
    }

    if (inputType === "class-support-note") {
      const classIndex = Number(target.dataset.classIndex);
      state.grading.classes[classIndex].classSupportNote = target.value;
      scheduleAutoSave();
    }
  });

  overallProgressNoteEl.addEventListener("input", () => {
    if (!state.grading) return;
    state.grading.overallProgressNote = overallProgressNoteEl.value;
    scheduleAutoSave();
  });

  if (globalSearchInputEl && globalSearchResultsEl && globalSearchWrapEl) {
    globalSearchInputEl.addEventListener("input", () => {
      renderGlobalSearch(globalSearchInputEl.value);
    });

    globalSearchInputEl.addEventListener("focus", () => {
      renderGlobalSearch(globalSearchInputEl.value);
    });

    globalSearchInputEl.addEventListener("keydown", (event) => {
      if (!globalSearchResultsEl || globalSearchResultsEl.classList.contains("hidden")) return;
      const hits = Array.from(globalSearchResultsEl.querySelectorAll("[data-role='search-hit']"));
      if (hits.length === 0) return;

      const key = event.key;
      if (key === "ArrowDown" || key === "Down") {
        event.preventDefault();
        const next = clamp(state.ui.globalSearch.activeIndex + 1, 0, hits.length - 1);
        setGlobalSearchActiveIndex(next);
        return;
      }

      if (key === "ArrowUp" || key === "Up") {
        event.preventDefault();
        const next = clamp(state.ui.globalSearch.activeIndex - 1, 0, hits.length - 1);
        setGlobalSearchActiveIndex(next);
        return;
      }

      if (key === "Enter") {
        event.preventDefault();
        const active = hits[state.ui.globalSearch.activeIndex] ?? hits[0];
        if (!(active instanceof HTMLElement)) return;
        const payload = active.dataset.payload;
        if (!payload) return;
        try {
          const hit = JSON.parse(decodeURIComponent(payload));
          applySearchSelection(hit);
          globalSearchInputEl.value = "";
          closeGlobalSearch();
          globalSearchInputEl.blur();
        } catch (error) {
        }
        return;
      }

      if (key === "Escape") {
        event.preventDefault();
        globalSearchInputEl.value = "";
        closeGlobalSearch();
        globalSearchInputEl.blur();
      }
    });

    globalSearchResultsEl.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) return;
      const hitBtn = target.closest("[data-role='search-hit']");
      if (!(hitBtn instanceof HTMLElement)) return;
      const payload = hitBtn.dataset.payload;
      if (!payload) return;
      try {
        const hit = JSON.parse(decodeURIComponent(payload));
        applySearchSelection(hit);
        globalSearchInputEl.value = "";
        closeGlobalSearch();
        globalSearchInputEl.blur();
      } catch (error) {
      }
    });

    document.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (globalSearchWrapEl.contains(target)) return;
      closeGlobalSearch();
    });
  }

  saveBtn.addEventListener("click", () => saveToLocal("Saved.", false));
  if (resetEnteredDataBtn) {
    resetEnteredDataBtn.addEventListener("click", async () => {
      if (!state.grading) {
        setSaveStatus("No grading data to reset yet.", "bad");
        return;
      }

      const confirmed = await confirmResetDataWarning();
      if (!confirmed) return;

      clearEnteredGradingData();
      renderGrader();
      setSaveStatus("Entered scores and notes were reset.", "good");
      scheduleAutoSave();
    });
  }
  exportBtn.addEventListener("click", exportToFile);

  importBtn.addEventListener("click", () => {
    importRequestedFromSetup = false;
    importFileInputEl.click();
  });

  if (setupImportBtn) {
    setupImportBtn.addEventListener("click", () => {
      importRequestedFromSetup = true;
      importFileInputEl.click();
    });
  }

  importFileInputEl.addEventListener("change", async () => {
    const file = importFileInputEl.files?.[0];
    await importFromFile(file);
    importFileInputEl.value = "";
    importRequestedFromSetup = false;
  });

  sheetExportClassSelectEl.addEventListener("change", () => {
    if (state.grading) {
      const classIndex = Number(sheetExportClassSelectEl.value);
      if (Number.isInteger(classIndex)) state.grading.activeClassIndex = classIndex;
      renderGrader();
      scheduleAutoSave();
      return;
    }
    refreshExportSelectors();
  });

  sheetExportStudentSelectEl.addEventListener("change", () => {
  refreshExportSelectors();
});

if (sheetExportModeSelectEl) {
  sheetExportModeSelectEl.addEventListener("change", () => {
    refreshExportSelectors();
  });
}

exportStudentSheetBtn.addEventListener("click", () => {
  if (!state.grading) return;
  const mode = exportStudentMode();
  const { classIndex, studentIndex } = selectedExportIndexes();

  if (classIndex === null) {
    setSaveStatus("Choose a class first.", "bad");
    return;
  }

  if (mode === "class") {
    exportStudentSheetsByClass(classIndex);
    return;
  }

  if (studentIndex === null) {
    setSaveStatus("Choose a student first.", "bad");
    return;
  }

  exportStudentSheet(classIndex, studentIndex);
});


  exportClassReportBtn.addEventListener("click", () => {
    if (!state.grading) return;
    const { classIndex } = selectedExportIndexes();
    if (classIndex === null) {
      setSaveStatus("Choose a class first.", "bad");
      return;
    }
    exportClassReport(classIndex);
  });

  overallHelpListEl.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.dataset.type === "open-help-student") {
      if (!state.grading) return;
      const classIndex = Number(target.dataset.classIndex);
      const studentIndex = Number(target.dataset.studentIndex);
      const classRecord = state.grading.classes[classIndex];
      if (!classRecord) return;
      if (!Number.isInteger(studentIndex) || studentIndex < 0 || studentIndex >= classRecord.students.length) return;
      state.grading.activeClassIndex = classIndex;
      classRecord.selectedStudentIndex = studentIndex;
      if (sheetExportClassSelectEl) sheetExportClassSelectEl.value = String(classIndex);
      collapseOverallSummary();
      renderGrader();
      scheduleAutoSave();
      return;
    }

    if (target.dataset.type !== "toggle-help-list") return;
    state.ui.showAllHelp = !state.ui.showAllHelp;
    refreshOverallMetrics();
  });

  classSnapshotGridEl.addEventListener("click", (event) => {
    if (!state.grading) return;
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;

    if (target.dataset.type === "add-class") {
      const className = nextClassName();
      const newClass = {
        name: className,
        classSupportNote: "",
        sectionOrder: state.grading.sections.map((_, sectionIndex) => sectionIndex),
        selectedStudentIndex: 0,
        students: [createBlankStudentRecord("Student 1")],
      };
      state.grading.classes.push(newClass);
      state.grading.activeClassIndex = state.grading.classes.length - 1;
      if (sheetExportClassSelectEl) {
        sheetExportClassSelectEl.value = String(state.grading.activeClassIndex);
      }
      renderGrader();
      setSaveStatus(`Added ${className}`, "good");
      scheduleAutoSave();
      return;
    }

    const card = target.closest("[data-type='select-class']");
    if (!(card instanceof HTMLElement)) return;
    const classIndex = Number(card.dataset.classIndex);
    if (!Number.isInteger(classIndex)) return;
    state.grading.activeClassIndex = classIndex;
    if (sheetExportClassSelectEl) sheetExportClassSelectEl.value = String(classIndex);
    renderGrader();
    scheduleAutoSave();
  });

  if (overallSummaryPanelEl) {
    overallSummaryPanelEl.addEventListener("toggle", () => {
      if (overallSummaryPanelEl.open) {
        clearTimeout(summaryCollapseTimer);
        overallSummaryPanelEl.classList.remove("closing");
      }
    });
  }

  window.addEventListener("beforeunload", () => {
    saveToLocal("Saved", true);
  });

  renderSetup();
  refreshExportSelectors();

  const loaded = loadFromLocal(true);
  if (!loaded) {
    showSetupView();
    setLastSavedText("");
    setSaveStatus("Auto-save keeps progress in this browser. Use Download Backup for another computer.");
  }
})();
