(function () {

  /* ============================================================
   * Apps Script API config
   * Paste the /exec URL of your deployed Apps Script web app here
   * after deploying (Deploy > New deployment > Web app).
   * ========================================================== */
  const API_URL = "https://script.google.com/macros/s/AKfycbx3mPttPjfu6O0G1QVAZUg0wV1yR-ssyTaWzXIrqq_F6DfLO1CPGtG46RzVCSLdpOeRDg/exec";

  /**
   * GET-style call to the Apps Script API — used for all read-only
   * lookups (dropdown data, address cascade, entry lookup).
   * `params` is a plain object of query-string parameters.
   */
  function apiGet(action, params) {
    const qs = new URLSearchParams(Object.assign({ action }, params || {}));
    return fetch(`${API_URL}?${qs.toString()}`)
      .then(res => {
        if (!res.ok) throw new Error("Network response was not OK (" + res.status + ")");
        return res.json();
      });
  }

  /**
   * POST-style call to the Apps Script API — used only for
   * submitForm, since it mutates the spreadsheet.
   *
   * Content-Type is deliberately "text/plain;charset=utf-8" (not
   * "application/json") so the browser treats this as a "simple
   * request" and skips the CORS preflight (OPTIONS) — Apps Script
   * web apps cannot respond to preflight requests.
   */
  function apiPost(action, payload) {
    return fetch(`${API_URL}?action=${encodeURIComponent(action)}`, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    }).then(res => {
      if (!res.ok) throw new Error("Network response was not OK (" + res.status + ")");
      return res.json();
    });
  }

  /**
   * Drop-in replacement for the old google.script.run bridge.
   * Routes each server function name to the right GET/POST call
   * and param shape, so the rest of this file barely changed.
   */
  function callApi(fnName, ...args) {
    switch (fnName) {
      case "getInitialFormData":
        return apiGet("getInitialFormData");
      case "getRegions":
        return apiGet("getRegions");
      case "getProvinces":
        return apiGet("getProvinces", { regionCode: args[0] });
      case "getCities":
        return apiGet("getCities", { provinceCode: args[0] });
      case "getCitiesForRegion":
        return apiGet("getCitiesForRegion", { regionCode: args[0] });
      case "getBarangays":
        return apiGet("getBarangays", { cityCode: args[0] });
      case "getPostalCode":
        return apiGet("getPostalCode", { cityName: args[0] });
      case "getEntryByReference":
        return apiGet("getEntryByReference", { email: args[0], code: args[1] });
      case "submitForm":
        return apiPost("submitForm", args[0]);
      default:
        return Promise.reject(new Error("Unknown API function: " + fnName));
    }
  }

  /* ============================================================
   * Element refs
   * ========================================================== */
  const el = id => document.getElementById(id);

  const regionEl = el("region"), provinceEl = el("province"), cityEl = el("city"), barangayEl = el("barangay");
  const presentAddressCountryEl = el("presentAddressCountry");
  const presentAddressPostalEl = el("presentAddressPostal");
  const regionOfBirthEl = el("regionOfBirth"), provinceOfBirthEl = el("provinceOfBirth");
  const countryOfBirthEl = el("countryOfBirth");
  const nationalityEl = el("nationality");
  const industryTypeEl = el("industryType");
  const branchClientEl = el("branchClient");
  const branchClientOtherEl = el("branchClientOther");
  const branchClientOtherWrapEl = el("branchClientOtherWrap");
  const useEmployerDefaultsEl = el("useEmployerDefaults");
  const employerFieldsEl = el("employerFields");
  const residenceContactNoEl = el("residenceContactNo");
  const tinEl = el("tin");
  const noTinEl = el("noTin");
  const sssGsisEl = el("sssGsis");
  const noSssGsisEl = el("noSssGsis");
  const expectedMonthlyTransactionEl = el("expectedMonthlyTransaction");

  const formEl = el("mainForm");
  const submitBtn = el("submitBtn"), spinner = el("submitSpinner"), clearBtn = el("clearBtn");
  const submitBtnLabelEl = el("submitBtnLabel");
  const banner = el("formBanner");
  const successCard = el("successCard"), submitAnotherBtn = el("submitAnotherBtn");
  const successTitleEl = el("successTitle"), successReferenceCodeEl = el("successReferenceCode"),
        successEmailEl = el("successEmail"), copyReferenceBtn = el("copyReferenceBtn");

  // "View/Edit my response" lookup
  const showLookupBtn = el("showLookupBtn"), lookupPanel = el("lookupPanel");
  const lookupEmailEl = el("lookupEmail"), lookupCodeEl = el("lookupCode");
  const lookupErrorEl = el("lookupError"), findEntryBtn = el("findEntryBtn"),
        lookupSpinner = el("lookupSpinner"), cancelLookupBtn = el("cancelLookupBtn");
  const editModeBanner = el("editModeBanner"), editModeCodeEl = el("editModeCode"),
        exitEditModeBtn = el("exitEditModeBtn");

  let isEditMode = false;
  let editReferenceEmail = "";
  let editReferenceCode = "";

  let provinceLevelExists = true;
  let birthProvinceLevelExists = true;
  let postalManuallyEdited = false;
  let validationRules = null; // populated from getInitialFormData()
  let employerDefaults = null; // populated from getInitialFormData()
  let othersValue = "OTHERS"; // populated from getInitialFormData()

  /* ============================================================
   * Generic helpers
   * ========================================================== */
  function fillSelect(selectEl, items, placeholder, preferredValue) {
    selectEl.innerHTML = "";
    const ph = document.createElement("option");
    ph.value = ""; ph.textContent = placeholder;
    selectEl.appendChild(ph);
    items.forEach(item => {
      const opt = document.createElement("option");
      const value = typeof item === "string" ? item : item.code;
      const label = typeof item === "string" ? item : item.name;
      opt.value = value;
      opt.textContent = label;
      opt.dataset.name = typeof item === "string" ? item : item.name;
      selectEl.appendChild(opt);
    });
    if (preferredValue) {
      const match = Array.from(selectEl.options).find(o => o.dataset.name === preferredValue || o.value === preferredValue);
      if (match) selectEl.value = match.value;
    }
  }

  function fillSimpleSelect(selectEl, stringItems, placeholder, preferredValue) {
    selectEl.innerHTML = "";
    const ph = document.createElement("option");
    ph.value = ""; ph.textContent = placeholder;
    selectEl.appendChild(ph);
    stringItems.forEach(name => {
      const opt = document.createElement("option");
      opt.value = name; opt.textContent = name; opt.dataset.name = name;
      selectEl.appendChild(opt);
    });
    if (preferredValue && stringItems.includes(preferredValue)) selectEl.value = preferredValue;
  }

  function resetSelect(selectEl, placeholder) {
    selectEl.innerHTML = "";
    const opt = document.createElement("option");
    opt.value = ""; opt.textContent = placeholder;
    selectEl.appendChild(opt);
    selectEl.disabled = true;
  }

  function showBanner(msg) { banner.textContent = msg; banner.classList.add("show"); }
  function hideBanner() { banner.classList.remove("show"); }

  function selectedName(selectEl) {
    const opt = selectEl.options[selectEl.selectedIndex];
    return opt ? (opt.dataset.name || "") : "";
  }

  function toTitleCase(str) {
    return String(str || "").toLowerCase().replace(/\b\p{L}/gu, ch => ch.toUpperCase());
  }

  // Promise wrapper used by the "retrieve my entry" rehydration
  // flow below, which needs to await one cascade level before
  // loading the next. Now just delegates to callApi().
  function gsRun(fnName, ...args) {
    return callApi(fnName, ...args);
  }

  /* ============================================================
   * Initial load — one round trip for everything
   * ========================================================== */
  callApi("getInitialFormData")
    .then(data => {

      fillSelect(regionEl, data.regions, "Select Region");
      fillSelect(regionOfBirthEl, data.regions, "Select Region");

      fillSimpleSelect(presentAddressCountryEl, data.standardValues.country, "Select Country", "Philippines");
      fillSimpleSelect(countryOfBirthEl, data.standardValues.country, "Select Country", "Philippines");
      fillSimpleSelect(nationalityEl, data.standardValues.nationality, "Select Nationality", "Filipino");
      fillSimpleSelect(industryTypeEl, data.standardValues.industry, "Select Industry Type");

      validationRules = data.validation;
      employerDefaults = data.employerDefaults || {};
      othersValue = data.othersValue || "OTHERS";

      // Branch & Client: sheet-driven list + an always-appended "Others" option.
      fillSimpleSelect(branchClientEl, data.standardValues.branchClient, "Select Branch & Client");
      const othersOpt = document.createElement("option");
      othersOpt.value = othersValue; othersOpt.textContent = "Others (specify)";
      branchClientEl.appendChild(othersOpt);

      // Branch & Client is intentionally left on its "Select Branch & Client"
      // placeholder when the form opens — it is never pre-selected from the
      // Employer Defaults sheet, so the user always has to choose it.
      updateBranchClientOtherVisibility();

      // Employer fields: pre-fill from live sheet defaults so they hold
      // real values whether or not "Use company defaults" stays checked.
      el("employerBusinessName").value = employerDefaults.businessName || "";
      fillSimpleSelect(industryTypeEl, data.standardValues.industry, "Select Industry Type", employerDefaults.industry || "");
      el("employerOfficeNo").value = employerDefaults.officeNo || "";
      el("employerBusinessAddress").value = employerDefaults.businessAddress || "";
      el("officeStreet").value = employerDefaults.officeStreet || "";
      el("officeCity").value = employerDefaults.officeCity || "";
      el("officeProvince").value = employerDefaults.officeProvince || "";
      el("officeCountry").value = employerDefaults.officeCountry || "Philippines";
      el("officePostal").value = employerDefaults.officePostal || "";

      // "Use company default employer information" is pre-checked in the
      // HTML; apply the resulting collapsed/disabled state now that the
      // fields actually have values in them.
      applyEmployerDefaultsToggle();

      // Placeholder text driven by server-configured rules, in case
      // Config.gs values are changed later.
      el("tin").placeholder = validationRules.tin.length + " digits";
      el("sssGsis").placeholder = validationRules.sssGsis.length + " digits";
      el("mobileNo").placeholder = validationRules.mobile.example;
      residenceContactNoEl.placeholder = validationRules.mobile.example;

    })
    .catch(err => {
      showBanner("Could not load form data. Please refresh the page.");
      console.error(err);
    });

  /* ============================================================
   * Address cascade: Region -> Province -> City -> Barangay
   * ========================================================== */
  function onRegionChange() {
    const regionCode = regionEl.value;
    resetSelect(provinceEl, "Loading provinces\u2026");
    resetSelect(cityEl, "Select Province first");
    resetSelect(barangayEl, "Select City/Municipality first");
    presentAddressPostalEl.value = "";
    postalManuallyEdited = false;

    if (!regionCode) { resetSelect(provinceEl, "Select Region first"); return; }

    callApi("getProvinces", regionCode)
      .then(provinces => {
        if (provinces && provinces.length) {
          provinceLevelExists = true;
          provinceEl.disabled = false;
          fillSelect(provinceEl, provinces, "Select Province");
        } else {
          provinceLevelExists = false;
          resetSelect(provinceEl, "N/A for this region");
          loadCitiesForRegion(regionCode);
        }
      })
      .catch(err => { showBanner("Could not load provinces."); console.error(err); });
  }

  function loadCitiesForRegion(regionCode) {
    resetSelect(cityEl, "Loading cities\u2026");
    callApi("getCitiesForRegion", regionCode)
      .then(cities => {
        if (cities && cities.length) { cityEl.disabled = false; fillSelect(cityEl, cities, "Select City/Municipality"); }
        else { resetSelect(cityEl, "No cities found \u2014 contact support"); }
      })
      .catch(err => { showBanner("Could not load cities."); console.error(err); });
  }

  function onProvinceChange() {
    const provinceCode = provinceEl.value;
    resetSelect(cityEl, "Loading cities\u2026");
    resetSelect(barangayEl, "Select City/Municipality first");
    presentAddressPostalEl.value = "";
    postalManuallyEdited = false;

    if (!provinceCode) { resetSelect(cityEl, "Select Province first"); return; }

    callApi("getCities", provinceCode)
      .then(cities => { cityEl.disabled = false; fillSelect(cityEl, cities, "Select City/Municipality"); })
      .catch(err => { showBanner("Could not load cities."); console.error(err); });
  }

  function onCityChange() {
    const cityCode = cityEl.value;
    resetSelect(barangayEl, "Loading barangays\u2026");
    postalManuallyEdited = false;

    if (!cityCode) { resetSelect(barangayEl, "Select City/Municipality first"); presentAddressPostalEl.value = ""; return; }

    callApi("getBarangays", cityCode)
      .then(barangays => { barangayEl.disabled = false; fillSelect(barangayEl, barangays, "Select Barangay"); })
      .catch(err => { showBanner("Could not load barangays."); console.error(err); });

    // Auto-fill postal code based on the selected city/municipality name
    const cityName = selectedName(cityEl);
    if (cityName) {
      presentAddressPostalEl.value = "";
      callApi("getPostalCode", cityName)
        .then(zip => {
          if (zip && !postalManuallyEdited) presentAddressPostalEl.value = zip;
        })
        .catch(() => { /* silent — field stays editable */ });
    }
  }

  regionEl.addEventListener("change", onRegionChange);
  provinceEl.addEventListener("change", onProvinceChange);
  cityEl.addEventListener("change", onCityChange);
  presentAddressPostalEl.addEventListener("input", () => { postalManuallyEdited = true; });

  /* ============================================================
   * Birth address helper cascade: Region Of Birth -> Province Of Birth
   * Reuses the same server-side getProvinces() used by the address
   * cascade above, so Province Of Birth stays accurate against real
   * PSGC data instead of being free-typed.
   * ========================================================== */
  function onRegionOfBirthChange() {
    const regionCode = regionOfBirthEl.value;
    resetSelect(provinceOfBirthEl, "Loading provinces\u2026");

    if (!regionCode) { resetSelect(provinceOfBirthEl, "Select Region Of Birth first"); return; }

    const regionName = selectedName(regionOfBirthEl);

    callApi("getProvinces", regionCode)
      .then(provinces => {
        if (provinces && provinces.length) {
          birthProvinceLevelExists = true;
          provinceOfBirthEl.disabled = false;
          fillSelect(provinceOfBirthEl, provinces, "Select Province");
        } else {
          // Province-less region (e.g. NCR): use the region itself as
          // the "province" value so the field still carries a value.
          birthProvinceLevelExists = false;
          fillSelect(provinceOfBirthEl, [{ code: regionCode, name: regionName }], "N/A for this region");
          provinceOfBirthEl.value = regionCode;
          provinceOfBirthEl.disabled = true;
        }
      })
      .catch(err => { showBanner("Could not load provinces for Region Of Birth."); console.error(err); });
  }

  regionOfBirthEl.addEventListener("change", onRegionOfBirthChange);

  /* ============================================================
   * Branch & Client "Others" toggle
   * ========================================================== */
  function updateBranchClientOtherVisibility() {
    const isOthers = branchClientEl.value === othersValue;
    branchClientOtherWrapEl.style.display = isOthers ? "" : "none";
    if (!isOthers) branchClientOtherEl.value = "";
  }

  branchClientEl.addEventListener("change", updateBranchClientOtherVisibility);

  /* ============================================================
   * Auto-capitalization & Uppercase
   * ========================================================== */
  ["firstname", "middlename", "lastname"].forEach(id => {
    el(id).addEventListener("blur", e => { e.target.value = toTitleCase(e.target.value); });
  });
  //occupationRank
  ["occupationRank"].forEach(id => {
    el(id).addEventListener("input", e => { e.target.value = e.target.value.toUpperCase(); });
  });
  

  /* ============================================================
   * Employer defaults toggle
   * ========================================================== */
  function applyEmployerDefaultsToggle() {
    const useDefaults = useEmployerDefaultsEl.checked;
    employerFieldsEl.classList.toggle("gf-collapsed", useDefaults);
    employerFieldsEl.querySelectorAll("input, select").forEach(f => { f.disabled = useDefaults; });
    // Occupation Rank is intentionally NOT part of employerFieldsEl, so it
    // always stays visible/enabled regardless of this toggle.
  }

  useEmployerDefaultsEl.addEventListener("change", applyEmployerDefaultsToggle);

  /* ============================================================
   * TIN / SSS-GSIS default value toggles
   * Rather than have the user type the dummy default digits by
   * hand, checking the box selects the default value for them and
   * locks the field so it can't be overtyped by mistake.
   * ========================================================== */
  function applyNoTinToggle() {
    if (noTinEl.checked) {
      tinEl.value = (validationRules && validationRules.tin.defaultValue) || "055555555";
      tinEl.disabled = true;
      markInvalid("tin", false);
    } else {
      tinEl.disabled = false;
      if (tinEl.value === ((validationRules && validationRules.tin.defaultValue) || "055555555")) {
        tinEl.value = "";
      }
    }
  }

  function applyNoSssGsisToggle() {
    if (noSssGsisEl.checked) {
      sssGsisEl.value = (validationRules && validationRules.sssGsis.defaultValue) || "0555555555";
      sssGsisEl.disabled = true;
      markInvalid("sssGsis", false);
    } else {
      sssGsisEl.disabled = false;
      if (sssGsisEl.value === ((validationRules && validationRules.sssGsis.defaultValue) || "0555555555")) {
        sssGsisEl.value = "";
      }
    }
  }

  noTinEl.addEventListener("change", applyNoTinToggle);
  noSssGsisEl.addEventListener("change", applyNoSssGsisToggle);

  /* ============================================================
   * v4.1: Retrieve & edit an existing response by email + code
   * ========================================================== */
  showLookupBtn.addEventListener("click", () => {
    const isOpen = lookupPanel.style.display !== "none";
    lookupPanel.style.display = isOpen ? "none" : "";
    if (!isOpen) lookupEmailEl.focus();
  });

  cancelLookupBtn.addEventListener("click", () => {
    lookupPanel.style.display = "none";
    lookupEmailEl.value = "";
    lookupCodeEl.value = "";
    lookupErrorEl.textContent = "";
    lookupErrorEl.classList.remove("show");
  });

  function setLookupLoading(isLoading) {
    findEntryBtn.disabled = isLoading;
    lookupSpinner.classList.toggle("show", isLoading);
  }

  function enterEditMode(email, code, hasFullSnapshot) {
    isEditMode = true;
    editReferenceEmail = email;
    editReferenceCode = code;
    editModeCodeEl.textContent = code;
    editModeBanner.style.display = "flex";
    editModeBanner.classList.add("show");
    submitBtnLabelEl.textContent = "Update My Information";

    if (!hasFullSnapshot) {
      editModeCodeEl.parentElement.append(
        Object.assign(document.createElement("span"), {
          textContent: " This entry predates the online edit feature, so please double-check every field (especially the address dropdowns) before submitting.",
        })
      );
    }
  }

  function exitEditMode() {
    isEditMode = false;
    editReferenceEmail = "";
    editReferenceCode = "";
    editModeBanner.style.display = "none";
    editModeBanner.classList.remove("show");
    submitBtnLabelEl.textContent = "Submit";
  }

  exitEditModeBtn.addEventListener("click", () => {
    formEl.reset();
    resetFormUI();
    exitEditMode();
    hideBanner();
  });

  findEntryBtn.addEventListener("click", async () => {
    const email = lookupEmailEl.value.trim();
    const code = lookupCodeEl.value.trim();

    lookupErrorEl.textContent = "";
    lookupErrorEl.classList.remove("show");

    if (!email || !code) {
      lookupErrorEl.textContent = "Please enter both your email and reference code.";
      lookupErrorEl.classList.add("show");
      return;
    }

    setLookupLoading(true);

    callApi("getEntryByReference", email, code)
      .then(async result => {
        setLookupLoading(false);

        if (!result || result.status !== "success") {
          lookupErrorEl.textContent = (result && result.message) || "No entry found for that email and reference code.";
          lookupErrorEl.classList.add("show");
          return;
        }

        try {
          await applyEntryToForm(result.data || {});
        } catch (e) {
          console.error(e);
        }

        enterEditMode(email, result.referenceCode, result.hasFullSnapshot);

        lookupPanel.style.display = "none";
        lookupEmailEl.value = "";
        lookupCodeEl.value = "";
        formEl.scrollIntoView({ behavior: "smooth", block: "start" });
      })
      .catch(err => {
        setLookupLoading(false);
        lookupErrorEl.textContent = "Something went wrong while looking up your entry. Please try again.";
        lookupErrorEl.classList.add("show");
        console.error(err);
      });
  });

  /**
   * Repopulates every field from a saved submission (the payload
   * shape produced by buildPayload()). Simple fields are set
   * directly; the Region/Province/City/Barangay cascades are async
   * (each level's options load from the server) so they're awaited
   * in sequence via restoreAddressCascade().
   */
  async function applyEntryToForm(data) {

    // Branch & Client
    if (data.branchClient) {
      const match = Array.from(branchClientEl.options).find(o => o.dataset.name === data.branchClient);
      if (match) {
        branchClientEl.value = match.value;
      } else {
        branchClientEl.value = othersValue;
        branchClientOtherEl.value = data.branchClient;
      }
      updateBranchClientOtherVisibility();
    }

    el("firstname").value = data.firstname || "";
    el("middlename").value = data.middlename || "";
    el("lastname").value = data.lastname || "";
    if (data.gender) el("gender").value = data.gender;
    if (data.maritalStatus) el("maritalStatus").value = data.maritalStatus;

    el("street").value = data.street || "";
    if (data.presentAddressCountry) presentAddressCountryEl.value = data.presentAddressCountry;

    // TIN / SSS-GSIS — restore via the "no TIN/SSS" toggle when the
    // saved value is the shared default, otherwise fill it directly.
    const tinDefault = (validationRules && validationRules.tin.defaultValue) || "055555555";
    const sssDefault = (validationRules && validationRules.sssGsis.defaultValue) || "0555555555";
    noTinEl.checked = data.tin === tinDefault;
    applyNoTinToggle();
    if (!noTinEl.checked) tinEl.value = data.tin || "";
    noSssGsisEl.checked = data.sssGsis === sssDefault;
    applyNoSssGsisToggle();
    if (!noSssGsisEl.checked) sssGsisEl.value = data.sssGsis || "";

    el("placeOfBirth").value = data.placeOfBirth || "";
    if (data.countryOfBirth) countryOfBirthEl.value = data.countryOfBirth;
    el("dateOfBirth").value = data.dateOfBirth || "";
    if (data.nationality) nationalityEl.value = data.nationality;

    el("mobileNo").value = data.mobileNo || "";
    residenceContactNoEl.value = data.residenceContactNo || "";
    el("email").value = data.email || "";

    useEmployerDefaultsEl.checked = !!data.useEmployerDefaults;
    el("employerBusinessName").value = data.employerBusinessName || "";
    if (data.industryType) industryTypeEl.value = data.industryType;
    el("employerOfficeNo").value = data.employerOfficeNo || "";
    el("occupationRank").value = data.occupationRank || "";
    el("employerBusinessAddress").value = data.employerBusinessAddress || "";
    el("officeStreet").value = data.officeStreet || "";
    el("officeCity").value = data.officeCity || "";
    el("officeProvince").value = data.officeProvince || "";
    el("officeCountry").value = data.officeCountry || "Philippines";
    el("officePostal").value = data.officePostal || "";
    applyEmployerDefaultsToggle();

    expectedMonthlyTransactionEl.value = data.expectedMonthlyTransaction || "";
    el("enableEBanking").checked = !!data.enableEBanking;
    el("availCreditCard").checked = !!data.availCreditCard;

    if (data.preferredMailingAddress) {
      const radio = document.querySelector(`input[name="preferredMailingAddress"][value="${data.preferredMailingAddress}"]`);
      if (radio) radio.checked = true;
    }

    document.querySelectorAll(".gf-invalid").forEach(f => f.classList.remove("gf-invalid"));
    document.querySelectorAll(".gf-error-text.show").forEach(f => f.classList.remove("show"));

    await restoreAddressCascade(data);
  }

  async function restoreAddressCascade(data) {

    // Present Address: Region -> Province -> City -> Barangay
    if (data.regionCode) {
      regionEl.value = data.regionCode;
      try {
        const provinces = await gsRun("getProvinces", data.regionCode);
        if (provinces && provinces.length) {
          provinceLevelExists = true;
          provinceEl.disabled = false;
          fillSelect(provinceEl, provinces, "Select Province");
          if (data.provinceCode) provinceEl.value = data.provinceCode;
        } else {
          provinceLevelExists = false;
          resetSelect(provinceEl, "N/A for this region");
        }

        const cities = (provinceLevelExists && data.provinceCode)
          ? await gsRun("getCities", data.provinceCode)
          : await gsRun("getCitiesForRegion", data.regionCode);

        if (cities && cities.length) {
          cityEl.disabled = false;
          fillSelect(cityEl, cities, "Select City/Municipality");
          if (data.cityCode) cityEl.value = data.cityCode;
        }

        if (data.cityCode) {
          const barangays = await gsRun("getBarangays", data.cityCode);
          barangayEl.disabled = false;
          fillSelect(barangayEl, barangays, "Select Barangay");
          if (data.barangayCode) barangayEl.value = data.barangayCode;
        }
      } catch (e) {
        console.error(e);
      }
    }

    presentAddressPostalEl.value = data.presentAddressPostal || "";
    postalManuallyEdited = true;

    // Region Of Birth -> Province Of Birth
    if (data.regionOfBirthCode) {
      regionOfBirthEl.value = data.regionOfBirthCode;
      try {
        const provinces = await gsRun("getProvinces", data.regionOfBirthCode);
        if (provinces && provinces.length) {
          birthProvinceLevelExists = true;
          provinceOfBirthEl.disabled = false;
          fillSelect(provinceOfBirthEl, provinces, "Select Province");
          if (data.provinceOfBirth) {
            const match = Array.from(provinceOfBirthEl.options).find(o => o.dataset.name === data.provinceOfBirth);
            if (match) provinceOfBirthEl.value = match.value;
          }
        } else {
          birthProvinceLevelExists = false;
          const regionName = selectedName(regionOfBirthEl);
          fillSelect(provinceOfBirthEl, [{ code: data.regionOfBirthCode, name: regionName }], "N/A for this region");
          provinceOfBirthEl.value = data.regionOfBirthCode;
          provinceOfBirthEl.disabled = true;
        }
      } catch (e) {
        console.error(e);
      }
    }
  }

  /* ============================================================
   * Validation
   * ========================================================== */
  function markInvalid(idOrEl, invalid) {
    const targetEl = typeof idOrEl === "string" ? el(idOrEl) : idOrEl;
    if (!targetEl) return;
    targetEl.classList && targetEl.classList.toggle("gf-invalid", invalid);
    const errKey = typeof idOrEl === "string" ? idOrEl : targetEl.id;
    const errEl = document.querySelector(`[data-error-for="${errKey}"]`);
    if (errEl) errEl.classList.toggle("show", invalid);
  }

  function validate() {
    let valid = true;

    const requiredText = ["firstname", "lastname", "street", "presentAddressPostal",
      "tin", "sssGsis", "placeOfBirth", "dateOfBirth", "mobileNo", "email", "occupationRank"];
    requiredText.forEach(id => {
      const bad = !el(id).value.trim();
      markInvalid(id, bad);
      if (bad) valid = false;
    });

    const requiredSelects = provinceLevelExists
      ? ["region", "province", "city", "barangay", "presentAddressCountry", "countryOfBirth", "nationality", "gender", "maritalStatus"]
      : ["region", "city", "barangay", "presentAddressCountry", "countryOfBirth", "nationality", "gender", "maritalStatus"];
    requiredSelects.forEach(id => {
      const bad = !el(id).value;
      markInvalid(id, bad);
      if (bad) valid = false;
    });

    // Region Of Birth / Province Of Birth
    const regionOfBirthBad = !regionOfBirthEl.value;
    markInvalid("regionOfBirth", regionOfBirthBad);
    if (regionOfBirthBad) valid = false;

    if (!regionOfBirthBad) {
      const provinceOfBirthBad = birthProvinceLevelExists && !provinceOfBirthEl.value;
      markInvalid("provinceOfBirth", provinceOfBirthBad);
      if (provinceOfBirthBad) valid = false;
    }

    // Branch & Client (+ "Others" free-text)
    const branchClientBad = !branchClientEl.value;
    markInvalid("branchClient", branchClientBad);
    if (branchClientBad) valid = false;

    if (branchClientEl.value === othersValue) {
      const branchClientOtherBad = !branchClientOtherEl.value.trim();
      markInvalid("branchClientOther", branchClientOtherBad);
      if (branchClientOtherBad) valid = false;
    }

    // Format checks
    if (validationRules) {
      if (!new RegExp(validationRules.mobile.pattern).test(el("mobileNo").value.trim())) { markInvalid("mobileNo", true); valid = false; }
      if (!new RegExp(validationRules.tin.pattern).test(el("tin").value.trim())) { markInvalid("tin", true); valid = false; }
      if (!new RegExp(validationRules.sssGsis.pattern).test(el("sssGsis").value.trim())) { markInvalid("sssGsis", true); valid = false; }
      if (!new RegExp(validationRules.postalCode.pattern).test(el("presentAddressPostal").value.trim())) { markInvalid("presentAddressPostal", true); valid = false; }

      const officePostalVal = el("officePostal").value.trim();
      if (!useEmployerDefaultsEl.checked && officePostalVal && !new RegExp(validationRules.postalCode.pattern).test(officePostalVal)) {
        markInvalid("officePostal", true); valid = false;
      }

      // Residence Contact No is optional, format-checked only if filled.
      const residenceVal = residenceContactNoEl.value.trim();
      if (residenceVal && !new RegExp(validationRules.mobile.pattern).test(residenceVal)) {
        markInvalid("residenceContactNo", true); valid = false;
      }

      // Expected Monthly Banking Transaction must be a number within range.
      const emtRange = validationRules.expectedMonthlyTransaction;
      const emtVal = Number(expectedMonthlyTransactionEl.value.trim());
      const emtBad = !expectedMonthlyTransactionEl.value.trim() || isNaN(emtVal) ||
        emtVal < emtRange.min || emtVal > emtRange.max;
      markInvalid("expectedMonthlyTransaction", emtBad);
      if (emtBad) valid = false;
    }

    // Preferred Mailing Address radio
    const mailingChosen = document.querySelector('input[name="preferredMailingAddress"]:checked');
    const mailingErrEl = document.querySelector('[data-error-for="preferredMailingAddress"]');
    if (!mailingChosen) {
      if (mailingErrEl) mailingErrEl.classList.add("show");
      valid = false;
    } else if (mailingErrEl) {
      mailingErrEl.classList.remove("show");
    }

    return valid;
  }

  /* ============================================================
   * Submit
   * ========================================================== */
  function buildPayload() {
    const mailingChosen = document.querySelector('input[name="preferredMailingAddress"]:checked');

    return {
      branchClient: branchClientEl.value,
      branchClientOther: branchClientOtherEl.value.trim(),
      firstname: toTitleCase(el("firstname").value),
      middlename: toTitleCase(el("middlename").value),
      lastname: toTitleCase(el("lastname").value),
      gender: el("gender").value,
      maritalStatus: el("maritalStatus").value,

      street: el("street").value.trim(),
      regionCode: regionEl.value,
      provinceCode: provinceEl.value,
      cityCode: cityEl.value,
      cityName: selectedName(cityEl),
      provinceName: provinceLevelExists ? selectedName(provinceEl) : "N/A",
      barangayCode: barangayEl.value,
      barangayName: selectedName(barangayEl),
      presentAddressCountry: presentAddressCountryEl.value,
      presentAddressPostal: presentAddressPostalEl.value.trim(),

      tin: el("tin").value.trim(),
      sssGsis: el("sssGsis").value.trim(),

      placeOfBirth: el("placeOfBirth").value.trim(),
      regionOfBirthCode: regionOfBirthEl.value,
      regionOfBirthName: selectedName(regionOfBirthEl),
      provinceOfBirth: birthProvinceLevelExists ? selectedName(provinceOfBirthEl) : selectedName(regionOfBirthEl),
      countryOfBirth: countryOfBirthEl.value,
      dateOfBirth: el("dateOfBirth").value, // yyyy-mm-dd from <input type=date>; reformatted server-side if needed
      nationality: nationalityEl.value,

      mobileNo: el("mobileNo").value.trim(),
      residenceContactNo: residenceContactNoEl.value.trim(),
      email: el("email").value.trim(),

      useEmployerDefaults: useEmployerDefaultsEl.checked,
      employerBusinessName: el("employerBusinessName").value.trim(),
      industryType: industryTypeEl.value,
      employerOfficeNo: el("employerOfficeNo").value.trim(),
      occupationRank: el("occupationRank").value.trim(),
      employerBusinessAddress: el("employerBusinessAddress").value.trim(),
      officeStreet: el("officeStreet").value.trim(),
      officeCity: el("officeCity").value.trim(),
      officeProvince: el("officeProvince").value.trim(),
      officeCountry: el("officeCountry").value.trim(),
      officePostal: el("officePostal").value.trim(),

      expectedMonthlyTransaction: el("expectedMonthlyTransaction").value.trim(),
      enableEBanking: el("enableEBanking").checked,
      availCreditCard: el("availCreditCard").checked,
      preferredMailingAddress: mailingChosen ? mailingChosen.value : "",

      referenceEmail: isEditMode ? editReferenceEmail : "",
      referenceCode: isEditMode ? editReferenceCode : ""
    };
  }

  function setSubmitting(isSubmitting) {
    submitBtn.disabled = isSubmitting;
    spinner.classList.toggle("show", isSubmitting);
  }

  submitBtn.addEventListener("click", () => {
    hideBanner();

    if (!validate()) {
      showBanner("Please fix the highlighted fields before submitting.");
      return;
    }

    setSubmitting(true);

    callApi("submitForm", buildPayload())
      .then(result => {
        setSubmitting(false);
        if (result.status === "success") {
          successTitleEl.textContent = result.message || "Your response has been recorded";
          successReferenceCodeEl.textContent = result.referenceCode || "";
          successEmailEl.textContent = result.email || "";
          exitEditMode();
          formEl.style.display = "none";
          successCard.classList.add("show");
        } else {
          showBanner(result.message || "Something went wrong. Please try again.");
        }
      })
      .catch(err => {
        setSubmitting(false);
        showBanner("Could not submit the form. Please check your connection and try again.");
        console.error(err);
      });
  });

  function resetFormUI() {
    resetSelect(provinceEl, "Select Region first");
    resetSelect(cityEl, "Select Province first");
    resetSelect(barangayEl, "Select City/Municipality first");
    presentAddressPostalEl.value = "";
    postalManuallyEdited = false;

    resetSelect(provinceOfBirthEl, "Select Region Of Birth first");

    branchClientOtherEl.value = "";
    branchClientEl.value = ""; // always back to "Select Branch & Client"
    updateBranchClientOtherVisibility();

    // formEl.reset() clears checked/value but not the disabled attribute,
    // so re-sync the TIN/SSS-GSIS fields with their (now unchecked) boxes.
    tinEl.disabled = false;
    sssGsisEl.disabled = false;

    useEmployerDefaultsEl.checked = true;
    if (employerDefaults) {
      el("employerBusinessName").value = employerDefaults.businessName || "";
      industryTypeEl.value = employerDefaults.industry || "";
      el("employerOfficeNo").value = employerDefaults.officeNo || "";
      el("employerBusinessAddress").value = employerDefaults.businessAddress || "";
      el("officeStreet").value = employerDefaults.officeStreet || "";
      el("officeCity").value = employerDefaults.officeCity || "";
      el("officeProvince").value = employerDefaults.officeProvince || "";
      el("officeCountry").value = employerDefaults.officeCountry || "Philippines";
      el("officePostal").value = employerDefaults.officePostal || "";
    }
    applyEmployerDefaultsToggle();

    document.querySelectorAll(".gf-invalid").forEach(f => f.classList.remove("gf-invalid"));
    document.querySelectorAll(".gf-error-text.show").forEach(f => f.classList.remove("show"));
  }

  clearBtn.addEventListener("click", () => {
    formEl.reset();
    resetFormUI();
    exitEditMode();
    hideBanner();
  });

  submitAnotherBtn.addEventListener("click", () => {
    formEl.reset();
    formEl.style.display = "block";
    successCard.classList.remove("show");
    resetFormUI();
    exitEditMode();
  });

  copyReferenceBtn.addEventListener("click", () => {
    const code = successReferenceCodeEl.textContent;
    if (!code || !navigator.clipboard || !navigator.clipboard.writeText) return;
    navigator.clipboard.writeText(code).then(() => {
      copyReferenceBtn.textContent = "Copied!";
      copyReferenceBtn.classList.add("copied");
      setTimeout(() => {
        copyReferenceBtn.textContent = "Copy";
        copyReferenceBtn.classList.remove("copied");
      }, 2000);
    });
  });

})();
