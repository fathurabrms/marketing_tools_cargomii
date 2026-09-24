/* ============================================================
   CARGOMII MARKETING TOOL
   APP.JS FULL — BAGIAN 1 / 3

   PENTING:
   - Hapus seluruh app.js lama
   - Paste BAGIAN 1
   - Sambung BAGIAN 2 tepat di bawahnya
   - Sambung BAGIAN 3 tepat di bawah BAGIAN 2
============================================================ */

"use strict";


/* ============================================================
   STORAGE KEYS
============================================================ */

const STORAGE_KEYS = {
    prospects: "cargomii_marketing_prospects_v3",
    profile: "cargomii_marketing_profile_v3"
};


/* ============================================================
   GLOBAL STATE
============================================================ */

let prospects = [];
let marketingProfile = {
    name: "",
    phone: "",
    target: 100
};

let excelRows = [];
let excelSelectedIds = new Set();
let selectedDataIds = new Set();

let deleteTargetId = null;
let deleteTargetIds = [];
let currentFollowUpId = null;
let toastTimer = null;


/* ============================================================
   SAFE DOM HELPER
============================================================ */

function el(id) {
    return document.getElementById(id);
}


/* ============================================================
   LOAD LOCAL STORAGE
============================================================ */

function loadStorage() {

    /* --------------------------------------------------------
       LOAD PROSPECTS
    -------------------------------------------------------- */

    try {

        const storedProspects =
            localStorage.getItem(
                STORAGE_KEYS.prospects
            );

        if (storedProspects) {

            const parsed =
                JSON.parse(storedProspects);

            if (Array.isArray(parsed)) {
                prospects = parsed;
            }

        }

    } catch (error) {

        console.error(
            "Gagal membaca data prospek:",
            error
        );

        prospects = [];

    }


    /* --------------------------------------------------------
       LOAD PROFILE
    -------------------------------------------------------- */

    try {

        const storedProfile =
            localStorage.getItem(
                STORAGE_KEYS.profile
            );

        if (storedProfile) {

            const parsed =
                JSON.parse(storedProfile);

            if (
                parsed &&
                typeof parsed === "object"
            ) {

                marketingProfile = {
                    name:
                        String(
                            parsed.name || ""
                        ),

                    phone:
                        String(
                            parsed.phone || ""
                        ),

                    target:
                        Number(
                            parsed.target || 100
                        )
                };

            }

        }

    } catch (error) {

        console.error(
            "Gagal membaca profil marketing:",
            error
        );

    }


    /* --------------------------------------------------------
       NORMALIZE OLD DATA
    -------------------------------------------------------- */

    prospects = prospects.map(
        normalizeProspect
    );

}


/* ============================================================
   SAVE PROSPECTS
============================================================ */

function saveProspects() {

    try {

        localStorage.setItem(
            STORAGE_KEYS.prospects,
            JSON.stringify(prospects)
        );

    } catch (error) {

        console.error(
            "Gagal menyimpan prospek:",
            error
        );

        showToast(
            "Data gagal disimpan di browser.",
            "error"
        );

    }

}


/* ============================================================
   SAVE MARKETING PROFILE
============================================================ */

function saveMarketingProfile() {

    try {

        localStorage.setItem(
            STORAGE_KEYS.profile,
            JSON.stringify(marketingProfile)
        );

    } catch (error) {

        console.error(
            "Gagal menyimpan profil:",
            error
        );

        showToast(
            "Profil gagal disimpan.",
            "error"
        );

    }

}


/* ============================================================
   GENERATE UNIQUE ID
============================================================ */

function createId(prefix = "data") {

    return (
        prefix +
        "_" +
        Date.now() +
        "_" +
        Math.random()
            .toString(36)
            .slice(2, 10)
    );

}


/* ============================================================
   ESCAPE HTML
============================================================ */

function escapeHTML(value) {

    return String(
        value ?? ""
    )
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");

}


/* ============================================================
   CLEAN TEXT
============================================================ */

function cleanText(value) {

    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    return String(value)
        .replace(/\s+/g, " ")
        .trim();

}


/* ============================================================
   CLEAN PHONE NUMBER

   Hasil utama aplikasi:
   081234567890

   Mendukung input:
   0812...
   62812...
   +62812...
   812...
============================================================ */

function normalizePhone(value) {

    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }

    let phone =
        String(value)
            .trim();


    /* --------------------------------------------------------
       EXCEL KADANG MEMBACA ANGKA SEPERTI:
       8.123456789E+10
    -------------------------------------------------------- */

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        phone =
            Math.trunc(value)
                .toString();

    }


    phone =
        phone.replace(/[^\d]/g, "");


    if (!phone) {
        return "";
    }


    /* --------------------------------------------------------
       +62 / 62 -> 0
    -------------------------------------------------------- */

    if (phone.startsWith("62")) {

        phone =
            "0" +
            phone.substring(2);

    }


    /* --------------------------------------------------------
       8123 -> 08123
    -------------------------------------------------------- */

    if (
        phone.startsWith("8")
    ) {

        phone =
            "0" + phone;

    }


    return phone;

}


/* ============================================================
   VALIDATE INDONESIAN MOBILE NUMBER

   Sesuai kebutuhan:
   hanya nomor yang berawalan 08.

   Range dibuat cukup longgar agar data bisnis tidak
   terlalu banyak terbuang.
============================================================ */

function isValidMobilePhone(value) {

    const phone =
        normalizePhone(value);

    return /^08\d{8,12}$/.test(
        phone
    );

}


/* ============================================================
   WHATSAPP INTERNATIONAL NUMBER
============================================================ */

function toWhatsAppNumber(value) {

    const phone =
        normalizePhone(value);

    if (!phone) {
        return "";
    }

    if (
        phone.startsWith("0")
    ) {

        return (
            "62" +
            phone.substring(1)
        );

    }

    return phone;

}


/* ============================================================
   FORMAT DISPLAY DATE
============================================================ */

function formatDateID(value) {

    if (!value) {
        return "-";
    }

    const date =
        new Date(value);

    if (
        Number.isNaN(
            date.getTime()
        )
    ) {
        return "-";
    }

    return new Intl.DateTimeFormat(
        "id-ID",
        {
            day: "2-digit",
            month: "2-digit",
            year: "numeric"
        }
    ).format(date);

}


/* ============================================================
   NORMALIZE PROSPECT

   Supaya data lama maupun data baru mempunyai struktur sama.
============================================================ */

function normalizeProspect(item = {}) {

    const now =
        new Date().toISOString();

    return {

        id:
            item.id ||
            createId("prospect"),

        company:
            cleanText(
                item.company ||
                item.name ||
                item.nama ||
                item.businessName
            ),

        phone:
            normalizePhone(
                item.phone ||
                item.whatsapp ||
                item.telephone ||
                item.mobile
            ),

        region:
            cleanText(
                item.region ||
                item.area ||
                item.wilayah
            ),

        pic:
            cleanText(
                item.pic ||
                item.contactPerson
            ),

        website:
            cleanText(
                item.website ||
                item.site
            ),

        source:
            cleanText(
                item.source ||
                "Manual"
            ),

        address:
            cleanText(
                item.address ||
                item.alamat
            ),

        notes:
            cleanText(
                item.notes ||
                item.note ||
                item.keterangan
            ),

        rating:
            cleanText(
                item.rating
            ),

        category:
            cleanText(
                item.category ||
                item.kategori
            ),

        maps:
            cleanText(
                item.maps ||
                item.googleMaps ||
                item.mapUrl
            ),

        status:
            cleanText(
                item.status ||
                "Belum Follow Up"
            ),

        activities: {

            sms:
                Boolean(
                    item.activities?.sms
                ),

            whatsapp:
                Boolean(
                    item.activities?.whatsapp
                ),

            call:
                Boolean(
                    item.activities?.call
                )

        },

        createdAt:
            item.createdAt ||
            item.date ||
            now,

        updatedAt:
            item.updatedAt ||
            now

    };

}


/* ============================================================
   GET PROSPECT
============================================================ */

function getProspectById(id) {

    return prospects.find(
        item =>
            String(item.id) ===
            String(id)
    ) || null;

}


/* ============================================================
   DUPLICATE CHECK

   Nomor HP dijadikan identifier utama.
============================================================ */

function prospectPhoneExists(
    phone,
    ignoreId = null
) {

    const normalized =
        normalizePhone(phone);

    if (!normalized) {
        return false;
    }

    return prospects.some(
        item => {

            if (
                ignoreId !== null &&
                String(item.id) ===
                String(ignoreId)
            ) {
                return false;
            }

            return (
                normalizePhone(
                    item.phone
                ) === normalized
            );

        }
    );

}


/* ============================================================
   SHOW TOAST
============================================================ */

function showToast(
    message,
    type = "success"
) {

    const toast =
        el("toast");

    const toastMessage =
        el("toastMessage");

    const iconWrapper =
        el("toastIconWrapper");

    if (
        !toast ||
        !toastMessage
    ) {

        console.log(
            `[${type}] ${message}`
        );

        return;

    }


    toastMessage.textContent =
        message;


    if (iconWrapper) {

        if (type === "error") {

            iconWrapper.innerHTML = `
                <i
                    data-lucide="circle-alert"
                    class="w-4 h-4 text-red-400"
                ></i>
            `;

        } else {

            iconWrapper.innerHTML = `
                <i
                    data-lucide="check-circle-2"
                    class="w-4 h-4 text-cargomii-green"
                ></i>
            `;

        }

    }


    toast.classList.remove(
        "hidden"
    );


    if (window.lucide) {
        lucide.createIcons();
    }


    if (toastTimer) {
        clearTimeout(toastTimer);
    }


    toastTimer =
        setTimeout(
            () => {

                toast.classList.add(
                    "hidden"
                );

            },
            2800
        );

}


/* ============================================================
   NAVIGATION
============================================================ */

function openTab(tabId) {

    const sections =
        document.querySelectorAll(
            ".tab-content"
        );

    sections.forEach(
        section => {

            section.classList.add(
                "hidden"
            );

        }
    );


    const target =
        el(tabId);

    if (target) {

        target.classList.remove(
            "hidden"
        );

    }


    const navButtons =
        document.querySelectorAll(
            ".nav-btn"
        );

    navButtons.forEach(
        button => {

            button.classList.toggle(
                "active",
                button.dataset.tab === tabId
            );

        }
    );


    const mobileNavigation =
        el("mobileNavigation");

    if (mobileNavigation) {

        mobileNavigation.value =
            tabId;

    }


    /* --------------------------------------------------------
       REFRESH ICONS
    -------------------------------------------------------- */

    if (window.lucide) {

        setTimeout(
            () => lucide.createIcons(),
            0
        );

    }

}


/* ============================================================
   SETUP NAVIGATION
============================================================ */

function setupNavigation() {

    document
        .querySelectorAll(
            ".nav-btn[data-tab]"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

                        openTab(
                            button.dataset.tab
                        );

                    }
                );

            }
        );


    document
        .querySelectorAll(
            "[data-go-tab]"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

                        openTab(
                            button.dataset.goTab
                        );

                    }
                );

            }
        );


    const mobileNavigation =
        el("mobileNavigation");

    if (mobileNavigation) {

        mobileNavigation.addEventListener(
            "change",
            () => {

                openTab(
                    mobileNavigation.value
                );

            }
        );

    }

}


/* ============================================================
   MARKETING PROFILE ELEMENTS
============================================================ */

const marketingProfileForm =
    el("marketingProfileForm");

const marketingNameInput =
    el("marketingNameInput");

const marketingPhoneInput =
    el("marketingPhoneInput");

const marketingTargetInput =
    el("marketingTargetInput");


/* ============================================================
   RENDER MARKETING PROFILE
============================================================ */

function renderMarketingProfile() {

    const name =
        marketingProfile.name.trim();

    const phone =
        marketingProfile.phone.trim();

    const target =
        Math.max(
            Number(
                marketingProfile.target || 0
            ),
            0
        );


    /* --------------------------------------------------------
       INPUTS
    -------------------------------------------------------- */

    if (
        marketingNameInput &&
        document.activeElement !==
            marketingNameInput
    ) {

        marketingNameInput.value =
            name;

    }


    if (
        marketingPhoneInput &&
        document.activeElement !==
            marketingPhoneInput
    ) {

        marketingPhoneInput.value =
            phone;

    }


    if (
        marketingTargetInput &&
        document.activeElement !==
            marketingTargetInput
    ) {

        marketingTargetInput.value =
            target || "";

    }


    /* --------------------------------------------------------
       SIDEBAR
    -------------------------------------------------------- */

    const sidebarName =
        el("sidebarMarketingName");

    if (sidebarName) {

        sidebarName.textContent =
            name || "Marketing";

    }


    const sidebarTarget =
        el("sidebarMarketingTarget");

    if (sidebarTarget) {

        sidebarTarget.textContent =
            `Target: ${target} Data`;

    }


    /* --------------------------------------------------------
       DASHBOARD
    -------------------------------------------------------- */

    const dashboardName =
        el("dashboardMarketingName");

    if (dashboardName) {

        dashboardName.textContent =
            name || "Marketing";

    }


    /* --------------------------------------------------------
       PROFILE PREVIEW
    -------------------------------------------------------- */

    const previewName =
        el("profilePreviewName");

    if (previewName) {

        previewName.textContent =
            name || "Marketing";

    }


    const previewPhone =
        el("profilePreviewPhone");

    if (previewPhone) {

        previewPhone.textContent =
            phone || "-";

    }


    const previewTarget =
        el("profilePreviewTarget");

    if (previewTarget) {

        previewTarget.textContent =
            target;

    }

}


/* ============================================================
   PROFILE SUBMIT
============================================================ */

function setupMarketingProfile() {

    if (!marketingProfileForm) {
        return;
    }


    marketingProfileForm.addEventListener(
        "submit",
        event => {

            event.preventDefault();


            const name =
                cleanText(
                    marketingNameInput?.value
                );

            const phone =
                normalizePhone(
                    marketingPhoneInput?.value
                );

            const target =
                Number(
                    marketingTargetInput?.value
                );


            /* ------------------------------------------------
               VALIDATION
            ------------------------------------------------ */

            if (!name) {

                showToast(
                    "Nama marketing wajib diisi.",
                    "error"
                );

                marketingNameInput?.focus();

                return;

            }


            if (!isValidMobilePhone(phone)) {

                showToast(
                    "Nomor marketing harus berupa nomor HP yang diawali 08.",
                    "error"
                );

                marketingPhoneInput?.focus();

                return;

            }


            if (
                !Number.isFinite(target) ||
                target < 1
            ) {

                showToast(
                    "Target marketing minimal 1 data.",
                    "error"
                );

                marketingTargetInput?.focus();

                return;

            }


            /* ------------------------------------------------
               SAVE
            ------------------------------------------------ */

            marketingProfile = {
                name,
                phone,
                target:
                    Math.floor(target)
            };


            saveMarketingProfile();

            renderMarketingProfile();
            updateMessageTemplate();

            renderDashboard();

            renderFollowUpProgressSafe();


            showToast(
                "Profil marketing berhasil disimpan."
            );

        }
    );

}


/* ============================================================
   DASHBOARD COUNTERS
============================================================ */

function countFollowUp() {

    return prospects.filter(
        item =>
            item.status === "Follow Up" ||
            item.status === "Tertarik" ||
            item.status === "Tidak Tertarik" ||
            item.activities?.sms ||
            item.activities?.whatsapp ||
            item.activities?.call
    ).length;

}


function countInterested() {

    return prospects.filter(
        item =>
            item.status === "Tertarik"
    ).length;

}


/* ============================================================
   RENDER DASHBOARD
============================================================ */

function renderDashboard() {

    const total =
        prospects.length;

    const target =
        Math.max(
            Number(
                marketingProfile.target || 0
            ),
            0
        );

    const followUp =
        countFollowUp();

    const interested =
        countInterested();

    const remaining =
        Math.max(
            target - total,
            0
        );

    const percentage =
        target > 0
            ? Math.min(
                Math.round(
                    (total / target) * 100
                ),
                100
            )
            : 0;


    /* --------------------------------------------------------
       STAT CARDS
    -------------------------------------------------------- */

    const statTarget =
        el("statTarget");

    if (statTarget) {
        statTarget.textContent =
            target;
    }


    const statData =
        el("statData");

    if (statData) {
        statData.textContent =
            total;
    }


    const statFollowUp =
        el("statFollowUp");

    if (statFollowUp) {
        statFollowUp.textContent =
            followUp;
    }


    const statInterested =
        el("statInterested");

    if (statInterested) {
        statInterested.textContent =
            interested;
    }


    /* --------------------------------------------------------
       PROGRESS
    -------------------------------------------------------- */

    const progressPercentage =
        el("progressPercentage");

    if (progressPercentage) {

        progressPercentage.textContent =
            `${percentage}%`;

    }


    const progressBar =
        el("progressBar");

    if (progressBar) {

        progressBar.style.width =
            `${percentage}%`;

    }


    const progressText =
        el("progressText");

    if (progressText) {

        progressText.textContent =
            `${total} / ${target}`;

    }


    /* --------------------------------------------------------
       SUMMARY
    -------------------------------------------------------- */

    const dashboardTotalData =
        el("dashboardTotalData");

    if (dashboardTotalData) {

        dashboardTotalData.textContent =
            total;

    }


    const dashboardRemaining =
        el("dashboardRemaining");

    if (dashboardRemaining) {

        dashboardRemaining.textContent =
            remaining;

    }


    const dashboardPotential =
        el("dashboardPotential");

    if (dashboardPotential) {

        dashboardPotential.textContent =
            interested;

    }

}


/* ============================================================
   SAFE FOLLOW UP PROGRESS CALL

   Fungsi aslinya dibuat pada BAGIAN 3.
============================================================ */

function renderFollowUpProgressSafe() {

    if (
        typeof renderFollowUpProgress ===
        "function"
    ) {

        renderFollowUpProgress();

    }

}


/* ============================================================
   DETECT REGION FROM ADDRESS

   Ini bukan geocoding/API.

   Sistem membaca kata wilayah dari alamat Excel.
============================================================ */

const REGION_PATTERNS = [

    "Jakarta Selatan",
    "Jakarta Timur",
    "Jakarta Barat",
    "Jakarta Utara",
    "Jakarta Pusat",

    "Tangerang Selatan",
    "Kabupaten Tangerang",
    "Kota Tangerang",

    "Kabupaten Bekasi",
    "Kota Bekasi",

    "Kabupaten Bogor",
    "Kota Bogor",

    "Kota Depok",

    "Bandung",
    "Surabaya",
    "Semarang",
    "Yogyakarta",
    "Solo",
    "Surakarta",
    "Malang",

    "Medan",
    "Palembang",
    "Pekanbaru",
    "Padang",
    "Batam",

    "Pontianak",
    "Banjarmasin",
    "Balikpapan",
    "Samarinda",

    "Makassar",
    "Manado",
    "Bitung",
    "Gorontalo",

    "Denpasar",
    "Mataram",
    "Kupang",

    "Jayapura",
    "Sorong",
    "Manokwari",

    "Jakarta",
    "Tangerang",
    "Bekasi",
    "Bogor",
    "Depok"

];


/* ============================================================
   DETECT REGION
============================================================ */

function detectRegionFromAddress(
    address
) {

    const text =
        cleanText(address)
            .toLowerCase();

    if (!text) {
        return "";
    }


    for (
        const region of REGION_PATTERNS
    ) {

        if (
            text.includes(
                region.toLowerCase()
            )
        ) {

            return region;

        }

    }


    /* --------------------------------------------------------
       FALLBACK:
       coba ambil kota/kabupaten dari teks.
    -------------------------------------------------------- */

    const cityMatch =
        cleanText(address).match(
            /\b(?:Kota|Kabupaten|Kab\.?)\s+([A-Za-zÀ-ÿ.' -]{3,40})/i
        );

    if (cityMatch) {

        return cleanText(
            cityMatch[0]
        );

    }


    return "";

}


/* ============================================================
   MANUAL FORM ELEMENTS
============================================================ */

const dataForm =
    el("dataForm");

const companyInput =
    el("companyInput");

const phoneInput =
    el("phoneInput");

const regionInput =
    el("regionInput");

const picInput =
    el("picInput");

const websiteInput =
    el("websiteInput");

const sourceInput =
    el("sourceInput");

const addressInput =
    el("addressInput");

const notesInput =
    el("notesInput");


/* ============================================================
   CREATE MANUAL PROSPECT
============================================================ */

function setupManualForm() {

    if (!dataForm) {
        return;
    }


    dataForm.addEventListener(
        "submit",
        event => {

            event.preventDefault();


            const company =
                cleanText(
                    companyInput?.value
                );

            const phone =
                normalizePhone(
                    phoneInput?.value
                );

            const address =
                cleanText(
                    addressInput?.value
                );

            let region =
                cleanText(
                    regionInput?.value
                );


            /* ------------------------------------------------
               VALIDATION
            ------------------------------------------------ */

            if (!company) {

                showToast(
                    "Nama perusahaan wajib diisi.",
                    "error"
                );

                companyInput?.focus();

                return;

            }


            if (!isValidMobilePhone(phone)) {

                showToast(
                    "Nomor harus nomor HP valid yang diawali 08.",
                    "error"
                );

                phoneInput?.focus();

                return;

            }


            /* ------------------------------------------------
               REGION AUTO DETECT
            ------------------------------------------------ */

            if (!region) {

                region =
                    detectRegionFromAddress(
                        address
                    );

            }


            /* ------------------------------------------------
               CREATE
            ------------------------------------------------ */

            const now =
                new Date().toISOString();


            const newProspect =
                normalizeProspect({

                    id:
                        createId(
                            "prospect"
                        ),

                    company,

                    phone,

                    region,

                    pic:
                        cleanText(
                            picInput?.value
                        ),

                    website:
                        cleanText(
                            websiteInput?.value
                        ),

                    source:
                        cleanText(
                            sourceInput?.value
                        ) || "Manual",

                    address,

                    notes:
                        cleanText(
                            notesInput?.value
                        ),

                    status:
                        "Belum Follow Up",

                    activities: {
                        sms: false,
                        whatsapp: false,
                        call: false
                    },

                    createdAt:
                        now,

                    updatedAt:
                        now

                });


            prospects.unshift(
                newProspect
            );


            saveProspects();

            dataForm.reset();


            if (sourceInput) {

                sourceInput.value =
                    "Manual";

            }


            refreshAll();


            showToast(
                `${company} berhasil ditambahkan.`
            );


            openTab(
                "data-saya"
            );

        }
    );

}


/* ============================================================
   PHONE INPUT AUTO CLEAN

   User boleh paste +62, nanti field otomatis diarahkan
   menjadi format 08 saat blur.
============================================================ */

function setupPhoneInputs() {

    [
        phoneInput,
        marketingPhoneInput
    ].forEach(
        input => {

            if (!input) {
                return;
            }


            input.addEventListener(
                "blur",
                () => {

                    const normalized =
                        normalizePhone(
                            input.value
                        );

                    if (normalized) {

                        input.value =
                            normalized;

                    }

                }
            );

        }
    );

}


/* ============================================================
   EXCEL ELEMENTS
============================================================ */

const excelDropZone =
    el("excelDropZone");

const excelFileInput =
    el("excelFileInput");

const chooseExcelBtn =
    el("chooseExcelBtn");

const excelFileInfo =
    el("excelFileInfo");

const excelFileName =
    el("excelFileName");

const excelFileSummary =
    el("excelFileSummary");

const removeExcelBtn =
    el("removeExcelBtn");

const excelPreviewSection =
    el("excelPreviewSection");

const excelPreviewTable =
    el("excelPreviewTable");

const excelPreviewEmpty =
    el("excelPreviewEmpty");

const excelSearchInput =
    el("excelSearchInput");

const selectAllExcelBtn =
    el("selectAllExcelBtn");

const unselectAllExcelBtn =
    el("unselectAllExcelBtn");

const excelMasterCheckbox =
    el("excelMasterCheckbox");

const excelTotalCount =
    el("excelTotalCount");

const excelSelectedCount =
    el("excelSelectedCount");

const excelBottomSelectedCount =
    el("excelBottomSelectedCount");

const importSelectedExcelBtn =
    el("importSelectedExcelBtn");


/* ============================================================
   EXCEL COLUMN ALIASES

   Dibuat fleksibel karena file hasil Google Maps scraper /
   Google Business bisa punya nama kolom berbeda-beda.
============================================================ */

const EXCEL_COLUMN_ALIASES = {

    company: [
        "name",
        "nama",
        "nama perusahaan",
        "company",
        "company name",
        "business",
        "business name",
        "title",
        "place name",
        "nama bisnis"
    ],

    phone: [
        "phone",
        "phone number",
        "telephone",
        "telephone number",
        "tel",
        "mobile",
        "mobile phone",
        "whatsapp",
        "wa",
        "nomor",
        "nomer",
        "nomor telepon",
        "nomer telepon",
        "no hp",
        "no. hp",
        "hp",
        "contact",
        "contact number"
    ],

    address: [
        "address",
        "alamat",
        "full address",
        "complete address",
        "formatted address",
        "street address"
    ],

    region: [
        "region",
        "wilayah",
        "area",
        "city",
        "kota",
        "kabupaten",
        "district",
        "location"
    ],

    website: [
        "website",
        "web",
        "site",
        "url",
        "domain"
    ],

    rating: [
        "rating",
        "stars",
        "star",
        "review rating",
        "google rating"
    ],

    category: [
        "category",
        "kategori",
        "type",
        "business category",
        "main category"
    ],

    maps: [
        "google maps",
        "google maps url",
        "maps",
        "maps url",
        "map url",
        "place url",
        "google map",
        "google map url",
        "link"
    ]

};


/* ============================================================
   NORMALIZE EXCEL HEADER
============================================================ */

function normalizeHeader(value) {

    return cleanText(value)
        .toLowerCase()
        .replace(/[_\-]+/g, " ")
        .replace(/[^\p{L}\p{N}\s.]/gu, "")
        .replace(/\s+/g, " ")
        .trim();

}


/* ============================================================
   FIND EXCEL VALUE

   Definisi lengkap proses import dilanjutkan pada BAGIAN 2.
============================================================ */

function findExcelValue(
    row,
    aliases
) {

    if (
        !row ||
        typeof row !== "object"
    ) {
        return "";
    }


    const keys =
        Object.keys(row);


    /* --------------------------------------------------------
       EXACT HEADER MATCH
    -------------------------------------------------------- */

    for (
        const alias of aliases
    ) {

        const normalizedAlias =
            normalizeHeader(alias);

        const exactKey =
            keys.find(
                key =>
                    normalizeHeader(key) ===
                    normalizedAlias
            );

        if (exactKey !== undefined) {

            const value =
                row[exactKey];

            if (
                value !== undefined &&
                value !== null &&
                cleanText(value) !== ""
            ) {

                return value;

            }

        }

    }


    /* --------------------------------------------------------
       PARTIAL HEADER MATCH
    -------------------------------------------------------- */

    for (
        const alias of aliases
    ) {

        const normalizedAlias =
            normalizeHeader(alias);

        const partialKey =
            keys.find(
                key => {

                    const normalizedKey =
                        normalizeHeader(key);

                    return (
                        normalizedKey.includes(
                            normalizedAlias
                        ) ||
                        normalizedAlias.includes(
                            normalizedKey
                        )
                    );

                }
            );


        if (
            partialKey !== undefined
        ) {

            const value =
                row[partialKey];

            if (
                value !== undefined &&
                value !== null &&
                cleanText(value) !== ""
            ) {

                return value;

            }

        }

    }


    return "";

}


/* ============================================================
   BAGIAN 1 SELESAI

   JANGAN TAMBAHKAN:
   });
   }
   </script>

   APP.JS BAGIAN 2 HARUS LANGSUNG DITEMPEL
   DI BAWAH BARIS INI.
============================================================ */

/* ============================================================
   CARGOMII MARKETING TOOL
   APP.JS FULL — BAGIAN 2 / 3

   LANJUTAN LANGSUNG DARI BAGIAN 1
============================================================ */


/* ============================================================
   EXTRACT PHONE FROM EXCEL VALUE

   Beberapa file bisa punya format:
   0812...
   +62 812...
   0812... / 021...
   0812..., 0853...
============================================================ */

function extractMobilePhone(value) {

    if (
        value === null ||
        value === undefined
    ) {
        return "";
    }


    /* --------------------------------------------------------
       JIKA EXCEL MEMBACA SEBAGAI ANGKA
    -------------------------------------------------------- */

    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        const normalized =
            normalizePhone(value);

        return isValidMobilePhone(
            normalized
        )
            ? normalized
            : "";

    }


    const text =
        String(value).trim();


    if (!text) {
        return "";
    }


    /* --------------------------------------------------------
       COBA NORMALIZE LANGSUNG
    -------------------------------------------------------- */

    const direct =
        normalizePhone(text);

    if (
        isValidMobilePhone(direct)
    ) {

        return direct;

    }


    /* --------------------------------------------------------
       CARI BEBERAPA NOMOR DI DALAM STRING

       Contoh:
       Telp: 021-xxx / WA: 0812-xxxx
    -------------------------------------------------------- */

    const candidates =
        text.match(
            /(?:\+?62|0)?8[\d\s().-]{7,16}\d/g
        ) || [];


    for (
        const candidate of candidates
    ) {

        const phone =
            normalizePhone(
                candidate
            );

        if (
            isValidMobilePhone(phone)
        ) {

            return phone;

        }

    }


    return "";

}


/* ============================================================
   EXTRACT WEBSITE
============================================================ */

function normalizeWebsite(value) {

    let website =
        cleanText(value);

    if (!website) {
        return "";
    }


    if (
        website.startsWith(
            "www."
        )
    ) {

        website =
            "https://" + website;

    }


    return website;

}


/* ============================================================
   NORMALIZE MAPS URL
============================================================ */

function normalizeMaps(value) {

    return cleanText(value);

}


/* ============================================================
   CONVERT RAW EXCEL ROW
============================================================ */

function convertExcelRow(
    rawRow,
    index
) {

    const companyValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.company
        );


    const phoneValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.phone
        );


    const addressValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.address
        );


    const regionValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.region
        );


    const websiteValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.website
        );


    const ratingValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.rating
        );


    const categoryValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.category
        );


    const mapsValue =
        findExcelValue(
            rawRow,
            EXCEL_COLUMN_ALIASES.maps
        );


    const company =
        cleanText(
            companyValue
        );


    const phone =
        extractMobilePhone(
            phoneValue
        );


    const address =
        cleanText(
            addressValue
        );


    let region =
        cleanText(
            regionValue
        );


    if (!region) {

        region =
            detectRegionFromAddress(
                address
            );

    }


    return {

        id:
            createId(
                `excel_${index}`
            ),

        company,

        phone,

        address,

        region,

        website:
            normalizeWebsite(
                websiteValue
            ),

        rating:
            cleanText(
                ratingValue
            ),

        category:
            cleanText(
                categoryValue
            ),

        maps:
            normalizeMaps(
                mapsValue
            ),

        source:
            "Excel / Google Business",

        raw:
            rawRow

    };

}


/* ============================================================
   DETECT POSSIBLE PHONE FROM ANY COLUMN

   Fallback jika nama kolom telepon di file tidak dikenali.
============================================================ */

function findPhoneAnywhere(
    rawRow
) {

    if (
        !rawRow ||
        typeof rawRow !== "object"
    ) {
        return "";
    }


    for (
        const value of Object.values(
            rawRow
        )
    ) {

        const phone =
            extractMobilePhone(
                value
            );

        if (phone) {

            return phone;

        }

    }


    return "";

}


/* ============================================================
   DETECT COMPANY FALLBACK
============================================================ */

function findCompanyFallback(
    rawRow
) {

    if (
        !rawRow ||
        typeof rawRow !== "object"
    ) {
        return "";
    }


    const keys =
        Object.keys(rawRow);


    /* --------------------------------------------------------
       COBA KOLOM PERTAMA YANG BUKAN NOMOR
    -------------------------------------------------------- */

    for (
        const key of keys
    ) {

        const value =
            cleanText(
                rawRow[key]
            );


        if (!value) {
            continue;
        }


        if (
            extractMobilePhone(value)
        ) {
            continue;
        }


        if (
            /^https?:\/\//i.test(value)
        ) {
            continue;
        }


        if (
            value.length >= 2 &&
            value.length <= 150
        ) {

            return value;

        }

    }


    return "Tanpa Nama";

}


/* ============================================================
   PARSE EXCEL ROWS
============================================================ */

function parseExcelRows(
    rawRows
) {

    if (
        !Array.isArray(rawRows)
    ) {
        return [];
    }


    const results = [];

    const seenPhones =
        new Set();


    rawRows.forEach(
        (
            rawRow,
            index
        ) => {

            if (
                !rawRow ||
                typeof rawRow !== "object"
            ) {
                return;
            }


            const converted =
                convertExcelRow(
                    rawRow,
                    index
                );


            /* ------------------------------------------------
               FALLBACK PHONE
            ------------------------------------------------ */

            if (!converted.phone) {

                converted.phone =
                    findPhoneAnywhere(
                        rawRow
                    );

            }


            /* ------------------------------------------------
               WAJIB NOMOR HP 08
            ------------------------------------------------ */

            if (
                !isValidMobilePhone(
                    converted.phone
                )
            ) {

                return;

            }


            /* ------------------------------------------------
               COMPANY FALLBACK
            ------------------------------------------------ */

            if (!converted.company) {

                converted.company =
                    findCompanyFallback(
                        rawRow
                    );

            }


            /* ------------------------------------------------
               REGION FALLBACK
            ------------------------------------------------ */

            if (!converted.region) {

                converted.region =
                    detectRegionFromAddress(
                        converted.address
                    );

            }


            /* ------------------------------------------------
               DUPLICATE DALAM FILE
            ------------------------------------------------ */

            const phoneKey =
                normalizePhone(
                    converted.phone
                );


            if (
                seenPhones.has(
                    phoneKey
                )
            ) {

                return;

            }


            seenPhones.add(
                phoneKey
            );


            results.push(
                converted
            );

        }
    );


    return results;

}


/* ============================================================
   READ EXCEL FILE
============================================================ */

async function readExcelFile(
    file
) {

    if (!file) {
        return;
    }


    /* --------------------------------------------------------
       CHECK XLSX LIBRARY
    -------------------------------------------------------- */

    if (
        typeof XLSX === "undefined"
    ) {

        showToast(
            "Library XLSX belum dimuat di index.html.",
            "error"
        );

        console.error(
            "XLSX is not defined."
        );

        return;

    }


    /* --------------------------------------------------------
       CHECK EXTENSION
    -------------------------------------------------------- */

    const fileName =
        String(
            file.name || ""
        );

    const validExtension =
        /\.(xlsx|xls)$/i.test(
            fileName
        );


    if (!validExtension) {

        showToast(
            "File harus berformat .xlsx atau .xls.",
            "error"
        );

        return;

    }


    try {

        const arrayBuffer =
            await file.arrayBuffer();


        const workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array",
                    cellDates: false,
                    raw: false
                }
            );


        if (
            !workbook.SheetNames ||
            workbook.SheetNames.length === 0
        ) {

            showToast(
                "Sheet Excel tidak ditemukan.",
                "error"
            );

            return;

        }


        /* ----------------------------------------------------
           AMBIL SHEET PERTAMA
        ---------------------------------------------------- */

        const firstSheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[
                firstSheetName
            ];


        const rawRows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: "",
                    raw: false
                }
            );


        /* ----------------------------------------------------
           PARSE
        ---------------------------------------------------- */

        excelRows =
            parseExcelRows(
                rawRows
            );


        /* ----------------------------------------------------
           DEFAULT: SEMUA DATA VALID DIPILIH
        ---------------------------------------------------- */

        excelSelectedIds =
            new Set(
                excelRows.map(
                    item => item.id
                )
            );


        /* ----------------------------------------------------
           FILE INFO
        ---------------------------------------------------- */

        if (excelFileInfo) {

            excelFileInfo.classList.remove(
                "hidden"
            );

        }


        if (excelFileName) {

            excelFileName.textContent =
                file.name;

        }


        if (excelFileSummary) {

            excelFileSummary.textContent =
                `${rawRows.length} baris dibaca • ${excelRows.length} nomor HP 08 ditemukan`;

        }


        /* ----------------------------------------------------
           PREVIEW
        ---------------------------------------------------- */

        if (excelPreviewSection) {

            excelPreviewSection.classList.remove(
                "hidden"
            );

        }


        if (excelSearchInput) {

            excelSearchInput.value =
                "";

        }


        renderExcelPreview();


        if (
            excelRows.length === 0
        ) {

            showToast(
                "Tidak ditemukan nomor HP valid berawalan 08.",
                "error"
            );

        } else {

            showToast(
                `${excelRows.length} data valid berhasil dibaca.`
            );

        }

    } catch (error) {

        console.error(
            "READ EXCEL ERROR:",
            error
        );


        resetExcelImport();


        showToast(
            "File Excel gagal dibaca.",
            "error"
        );

    }

}


/* ============================================================
   RESET EXCEL IMPORT
============================================================ */

function resetExcelImport() {

    excelRows = [];

    excelSelectedIds.clear();


    if (excelFileInput) {

        excelFileInput.value =
            "";

    }


    if (excelFileInfo) {

        excelFileInfo.classList.add(
            "hidden"
        );

    }


    if (excelPreviewSection) {

        excelPreviewSection.classList.add(
            "hidden"
        );

    }


    if (excelFileName) {

        excelFileName.textContent =
            "-";

    }


    if (excelFileSummary) {

        excelFileSummary.textContent =
            "0 data ditemukan";

    }


    if (excelSearchInput) {

        excelSearchInput.value =
            "";

    }


    updateExcelCounters();

}


/* ============================================================
   SETUP EXCEL UPLOAD
============================================================ */

function setupExcelUpload() {

    /* --------------------------------------------------------
       BUTTON PILIH FILE
    -------------------------------------------------------- */

    if (
        chooseExcelBtn &&
        excelFileInput
    ) {

        chooseExcelBtn.addEventListener(
            "click",
            event => {

                event.stopPropagation();

                excelFileInput.click();

            }
        );

    }


    /* --------------------------------------------------------
       CLICK DROPZONE
    -------------------------------------------------------- */

    if (
        excelDropZone &&
        excelFileInput
    ) {

        excelDropZone.addEventListener(
            "click",
            event => {

                if (
                    event.target.closest(
                        "#chooseExcelBtn"
                    )
                ) {
                    return;
                }

                excelFileInput.click();

            }
        );

    }


    /* --------------------------------------------------------
       FILE CHANGE
    -------------------------------------------------------- */

    if (excelFileInput) {

        excelFileInput.addEventListener(
            "change",
            event => {

                const file =
                    event.target.files?.[0];

                if (file) {

                    readExcelFile(
                        file
                    );

                }

            }
        );

    }


    /* --------------------------------------------------------
       DRAG OVER
    -------------------------------------------------------- */

    if (excelDropZone) {

        excelDropZone.addEventListener(
            "dragover",
            event => {

                event.preventDefault();

                excelDropZone.classList.add(
                    "border-cargomii-blue",
                    "bg-blue-50/30"
                );

            }
        );


        excelDropZone.addEventListener(
            "dragleave",
            event => {

                event.preventDefault();

                excelDropZone.classList.remove(
                    "border-cargomii-blue",
                    "bg-blue-50/30"
                );

            }
        );


        excelDropZone.addEventListener(
            "drop",
            event => {

                event.preventDefault();

                excelDropZone.classList.remove(
                    "border-cargomii-blue",
                    "bg-blue-50/30"
                );


                const file =
                    event.dataTransfer
                        ?.files?.[0];


                if (file) {

                    readExcelFile(
                        file
                    );

                }

            }
        );

    }


    /* --------------------------------------------------------
       REMOVE FILE
    -------------------------------------------------------- */

    if (removeExcelBtn) {

        removeExcelBtn.addEventListener(
            "click",
            event => {

                event.stopPropagation();

                resetExcelImport();

                showToast(
                    "File Excel dihapus dari preview."
                );

            }
        );

    }


    /* --------------------------------------------------------
       SEARCH
    -------------------------------------------------------- */

    if (excelSearchInput) {

        excelSearchInput.addEventListener(
            "input",
            renderExcelPreview
        );

    }


    /* --------------------------------------------------------
       SELECT ALL
    -------------------------------------------------------- */

    if (selectAllExcelBtn) {

        selectAllExcelBtn.addEventListener(
            "click",
            () => {

                const visibleRows =
                    getFilteredExcelRows();


                visibleRows.forEach(
                    item => {

                        if (!findDuplicateProspect(item)) {
                            excelSelectedIds.add(
                                item.id
                            );
                        }

                    }
                );


                renderExcelPreview();

            }
        );

    }


    /* --------------------------------------------------------
       UNSELECT ALL
    -------------------------------------------------------- */

    if (unselectAllExcelBtn) {

        unselectAllExcelBtn.addEventListener(
            "click",
            () => {

                excelSelectedIds.clear();

                renderExcelPreview();

            }
        );

    }


    /* --------------------------------------------------------
       MASTER CHECKBOX
    -------------------------------------------------------- */

    if (excelMasterCheckbox) {

        excelMasterCheckbox.addEventListener(
            "change",
            () => {

                const visibleRows =
                    getFilteredExcelRows();


                if (
                    excelMasterCheckbox.checked
                ) {

                    visibleRows.forEach(
                        item => {

                            if (!findDuplicateProspect(item)) {
                                excelSelectedIds.add(
                                    item.id
                                );
                            }

                        }
                    );

                } else {

                    visibleRows.forEach(
                        item => {

                            excelSelectedIds.delete(
                                item.id
                            );

                        }
                    );

                }


                renderExcelPreview();

            }
        );

    }


    /* --------------------------------------------------------
       IMPORT
    -------------------------------------------------------- */

    if (importSelectedExcelBtn) {

        importSelectedExcelBtn.addEventListener(
            "click",
            importSelectedExcelRows
        );

    }

}


/* ============================================================
   GET FILTERED EXCEL ROWS
============================================================ */

function getFilteredExcelRows() {

    const keyword =
        cleanText(
            excelSearchInput?.value
        )
            .toLowerCase();


    if (!keyword) {

        return excelRows;

    }


    return excelRows.filter(
        item => {

            const searchable =
                [
                    item.company,
                    item.phone,
                    item.region,
                    item.address,
                    item.website,
                    item.category,
                    item.rating
                ]
                    .join(" ")
                    .toLowerCase();


            return searchable.includes(
                keyword
            );

        }
    );

}


/* ============================================================
   EXCEL MAPS LINK
============================================================ */

function getMapsCellHTML(
    item
) {

    const maps =
        cleanText(
            item.maps
        );


    if (!maps) {

        return `
            <span class="text-slate-300">
                -
            </span>
        `;

    }


    if (
        /^https?:\/\//i.test(
            maps
        )
    ) {

        return `
            <a
                href="${escapeHTML(maps)}"
                target="_blank"
                rel="noopener noreferrer"
                class="
                    inline-flex
                    items-center
                    gap-1.5
                    text-[9px]
                    font-bold
                    text-cargomii-blue
                    hover:underline
                "
            >
                <i
                    data-lucide="map-pin"
                    class="w-3 h-3"
                ></i>

                Maps
            </a>
        `;

    }


    return `
        <span class="text-[9px] text-slate-500">
            ${escapeHTML(maps)}
        </span>
    `;

}


/* ============================================================
   RENDER EXCEL PREVIEW
============================================================ */

function renderExcelPreview() {

    if (!excelPreviewTable) {
        return;
    }


    const rows =
        getFilteredExcelRows();


    excelPreviewTable.innerHTML =
        "";


    /* --------------------------------------------------------
       EMPTY
    -------------------------------------------------------- */

    if (
        rows.length === 0
    ) {

        if (excelPreviewEmpty) {

            excelPreviewEmpty.classList.remove(
                "hidden"
            );

        }

    } else {

        if (excelPreviewEmpty) {

            excelPreviewEmpty.classList.add(
                "hidden"
            );

        }

    }


    /* --------------------------------------------------------
       ROWS
    -------------------------------------------------------- */

    rows.forEach(
        (
            item,
            index
        ) => {

            const selected =
                excelSelectedIds.has(
                    item.id
                );

            const duplicateProspect =
                findDuplicateProspect(item);

            if (duplicateProspect) {
                excelSelectedIds.delete(item.id);
            }


            const tr =
                document.createElement(
                    "tr"
                );


            tr.className = duplicateProspect
                ? "border-b border-red-100 bg-red-50/60"
                : "border-b border-slate-100 hover:bg-slate-50/70 transition";


            tr.innerHTML = `

                <td class="p-3 align-top">

                    <input
                        type="checkbox"
                        class="
                            excel-row-checkbox
                            w-4 h-4
                            accent-blue-600
                        "
                        data-id="${escapeHTML(item.id)}"
                        ${selected && !duplicateProspect ? "checked" : ""}
                        ${duplicateProspect ? "disabled" : ""}
                        title="${duplicateProspect ? "Data ini sudah ada di Data Saya" : "Pilih data"}"
                    >

                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-400
                        align-top
                    "
                >
                    ${index + 1}
                </td>


                <td class="p-3 align-top">

                    <p
                        class="
                            text-[10px]
                            font-bold
                            text-cargomii-navy
                            max-w-[220px]
                        "
                    >
                        ${escapeHTML(
                            item.company || "-"
                        )}
                    </p>

                    ${
                        duplicateProspect
                            ? `
                                <span class="inline-flex mt-1.5 px-2 py-1 rounded-md bg-red-100 text-red-600 text-[8px] font-bold">
                                    Sudah ada di Data Saya
                                </span>
                              `
                            : ""
                    }

                    ${
                        item.region
                            ? `
                                <p
                                    class="
                                        text-[8px]
                                        text-slate-400
                                        mt-1
                                    "
                                >
                                    ${escapeHTML(item.region)}
                                </p>
                              `
                            : ""
                    }

                </td>


                <td class="p-3 align-top">

                    <span
                        class="
                            inline-flex
                            px-2.5 py-1.5
                            rounded-lg
                            bg-green-50
                            text-green-700
                            text-[9px]
                            font-bold
                            whitespace-nowrap
                        "
                    >
                        ${escapeHTML(item.phone)}
                    </span>

                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        leading-relaxed
                        text-slate-500
                        align-top
                        max-w-[280px]
                    "
                >
                    ${escapeHTML(
                        item.address || "-"
                    )}
                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                        max-w-[180px]
                    "
                >

                    ${
                        item.website
                            ? `
                                <span class="break-all">
                                    ${escapeHTML(item.website)}
                                </span>
                              `
                            : "-"
                    }

                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                    "
                >
                    ${escapeHTML(
                        item.rating || "-"
                    )}
                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                        max-w-[180px]
                    "
                >
                    ${escapeHTML(
                        item.category || "-"
                    )}
                </td>


                <td class="p-3 align-top">
                    ${getMapsCellHTML(item)}
                </td>

            `;


            excelPreviewTable.appendChild(
                tr
            );

        }
    );


    /* --------------------------------------------------------
       CHECKBOX EVENTS
    -------------------------------------------------------- */

    excelPreviewTable
        .querySelectorAll(
            ".excel-row-checkbox"
        )
        .forEach(
            checkbox => {

                checkbox.addEventListener(
                    "change",
                    () => {

                        const id =
                            checkbox.dataset.id;


                        if (
                            checkbox.checked
                        ) {

                            excelSelectedIds.add(
                                id
                            );

                        } else {

                            excelSelectedIds.delete(
                                id
                            );

                        }


                        updateExcelCounters();

                    }
                );

            }
        );


    updateExcelCounters();


    if (window.lucide) {

        lucide.createIcons();

    }

}


/* ============================================================
   UPDATE EXCEL COUNTERS
============================================================ */

function updateExcelCounters() {

    const total =
        excelRows.length;

    const selected =
        excelSelectedIds.size;


    if (excelTotalCount) {

        excelTotalCount.textContent =
            total;

    }


    if (excelSelectedCount) {

        excelSelectedCount.textContent =
            selected;

    }


    if (excelBottomSelectedCount) {

        excelBottomSelectedCount.textContent =
            selected;

    }


    if (importSelectedExcelBtn) {

        importSelectedExcelBtn.disabled =
            selected === 0;

    }


    /* --------------------------------------------------------
       MASTER CHECKBOX STATE BERDASARKAN ROW YANG TERLIHAT
    -------------------------------------------------------- */

    if (excelMasterCheckbox) {

        const visibleRows =
            getFilteredExcelRows().filter(
                item => !findDuplicateProspect(item)
            );


        if (
            visibleRows.length === 0
        ) {

            excelMasterCheckbox.checked =
                false;

            excelMasterCheckbox.indeterminate =
                false;

        } else {

            const selectedVisible =
                visibleRows.filter(
                    item =>
                        excelSelectedIds.has(
                            item.id
                        )
                ).length;


            excelMasterCheckbox.checked =
                selectedVisible ===
                visibleRows.length;


            excelMasterCheckbox.indeterminate =
                selectedVisible > 0 &&
                selectedVisible <
                    visibleRows.length;

        }

    }

}


/* ============================================================
   IMPORT SELECTED EXCEL DATA
============================================================ */

function importSelectedExcelRows() {

    const selectedRows =
        excelRows.filter(
            item =>
                excelSelectedIds.has(
                    item.id
                )
        );


    if (
        selectedRows.length === 0
    ) {

        showToast(
            "Pilih minimal satu data.",
            "error"
        );

        return;

    }


    let imported = 0;
    let duplicate = 0;


    selectedRows.forEach(
        item => {

            const phone =
                normalizePhone(
                    item.phone
                );

            const region =
                item.region ||
                detectRegionFromAddress(
                    item.address
                );


            if (
                !isValidMobilePhone(phone)
            ) {
                return;
            }


            /* ------------------------------------------------
               SKIP DUPLICATE DATA SAYA
            ------------------------------------------------ */

            if (
                findDuplicateProspect({
                    company: item.company,
                    phone,
                    region
                })
            ) {

                duplicate++;

                return;

            }


            const now =
                new Date().toISOString();


            const prospect =
                normalizeProspect({

                    id:
                        createId(
                            "prospect"
                        ),

                    company:
                        item.company ||
                        "Tanpa Nama",

                    phone,

                    region,

                    pic:
                        "",

                    website:
                        item.website,

                    source:
                        "Excel / Google Business",

                    address:
                        item.address,

                    notes:
                        "",

                    rating:
                        item.rating,

                    category:
                        item.category,

                    maps:
                        item.maps,

                    status:
                        "Belum Follow Up",

                    activities: {
                        sms: false,
                        whatsapp: false,
                        call: false
                    },

                    createdAt:
                        now,

                    updatedAt:
                        now

                });


            prospects.push(
                prospect
            );


            imported++;

        }
    );


    if (
        imported === 0
    ) {

        if (duplicate > 0) {

            showToast(
                `${duplicate} data sudah ada di Data Saya.`,
                "error"
            );

        } else {

            showToast(
                "Tidak ada data yang dapat dimasukkan.",
                "error"
            );

        }

        return;

    }


    saveProspects();

    refreshAll();


    if (duplicate > 0) {

        showToast(
            `${imported} data masuk, ${duplicate} data duplikat dilewati.`
        );

    } else {

        showToast(
            `${imported} data berhasil dimasukkan ke Data Saya.`
        );

    }


    resetExcelImport();

    openTab(
        "data-saya"
    );

}


/* ============================================================
   DATA SAYA ELEMENTS
============================================================ */

const dataTable =
    el("dataTable");

const dataEmptyState =
    el("dataEmptyState");

const searchInput =
    el("searchInput");

const statusFilter =
    el("statusFilter");

const dataSummaryTotal =
    el("dataSummaryTotal");

const dataSummaryPending =
    el("dataSummaryPending");

const dataSummaryFollowUp =
    el("dataSummaryFollowUp");

const dataSummaryInterested =
    el("dataSummaryInterested");

const deleteModal =
    el("deleteModal");

const cancelDeleteBtn =
    el("cancelDeleteBtn");

const confirmDeleteBtn =
    el("confirmDeleteBtn");

const selectAllDataCheckbox =
    el("selectAllDataCheckbox");

const deleteSelectedDataBtn =
    el("deleteSelectedDataBtn");

const deleteSelectedDataText =
    el("deleteSelectedDataText");


/* ============================================================
   GET FILTERED PROSPECTS
============================================================ */

function getFilteredProspects() {

    const keyword =
        cleanText(
            searchInput?.value
        )
            .toLowerCase();


    const selectedStatus =
        statusFilter?.value ||
        "all";


    return prospects.filter(
        item => {

            /* ------------------------------------------------
               SEARCH
            ------------------------------------------------ */

            const searchable =
                [
                    item.company,
                    item.phone,
                    item.region,
                    item.pic,
                    item.source,
                    item.address,
                    item.website,
                    item.category,
                    item.notes
                ]
                    .join(" ")
                    .toLowerCase();


            const matchSearch =
                !keyword ||
                searchable.includes(
                    keyword
                );


            /* ------------------------------------------------
               STATUS
            ------------------------------------------------ */

            const matchStatus =
                selectedStatus === "all" ||
                item.status ===
                    selectedStatus;


            return (
                matchSearch &&
                matchStatus
            );

        }
    );

}


/* ============================================================
   STATUS BADGE
============================================================ */

function getStatusBadge(
    status
) {

    switch (status) {

        case "Tertarik":

            return `
                <span
                    class="
                        inline-flex
                        px-2.5 py-1.5
                        rounded-lg
                        bg-green-50
                        text-green-700
                        text-[8px]
                        font-bold
                        whitespace-nowrap
                    "
                >
                    Tertarik
                </span>
            `;


        case "Tidak Tertarik":

            return `
                <span
                    class="
                        inline-flex
                        px-2.5 py-1.5
                        rounded-lg
                        bg-red-50
                        text-red-600
                        text-[8px]
                        font-bold
                        whitespace-nowrap
                    "
                >
                    Tidak Tertarik
                </span>
            `;


        case "Follow Up":

            return `
                <span
                    class="
                        inline-flex
                        px-2.5 py-1.5
                        rounded-lg
                        bg-amber-50
                        text-amber-700
                        text-[8px]
                        font-bold
                        whitespace-nowrap
                    "
                >
                    Follow Up
                </span>
            `;


        default:

            return `
                <span
                    class="
                        inline-flex
                        px-2.5 py-1.5
                        rounded-lg
                        bg-slate-100
                        text-slate-500
                        text-[8px]
                        font-bold
                        whitespace-nowrap
                    "
                >
                    Belum Follow Up
                </span>
            `;

    }

}


/* ============================================================
   ACTIVITY BADGES
============================================================ */

function getActivityBadges(
    item
) {

    const activities =
        item.activities || {};


    const smsClass =
        activities.sms
            ? "bg-blue-50 text-blue-600"
            : "bg-slate-50 text-slate-300";


    const waClass =
        activities.whatsapp
            ? "bg-green-50 text-green-600"
            : "bg-slate-50 text-slate-300";


    const callClass =
        activities.call
            ? "bg-violet-50 text-violet-600"
            : "bg-slate-50 text-slate-300";


    return `

        <div class="flex items-center gap-1">

            <span
                title="SMS"
                class="
                    w-7 h-7
                    rounded-lg
                    flex
                    items-center
                    justify-center
                    ${smsClass}
                "
            >
                <i
                    data-lucide="message-square-text"
                    class="w-3 h-3"
                ></i>
            </span>


            <span
                title="WhatsApp"
                class="
                    w-7 h-7
                    rounded-lg
                    flex
                    items-center
                    justify-center
                    ${waClass}
                "
            >
                <i
                    data-lucide="message-circle"
                    class="w-3 h-3"
                ></i>
            </span>


            <span
                title="Call"
                class="
                    w-7 h-7
                    rounded-lg
                    flex
                    items-center
                    justify-center
                    ${callClass}
                "
            >
                <i
                    data-lucide="phone"
                    class="w-3 h-3"
                ></i>
            </span>

        </div>

    `;

}


/* ============================================================
   RENDER DATA SUMMARY
============================================================ */

function renderDataSummary() {

    const total =
        prospects.length;


    const pending =
        prospects.filter(
            item =>
                item.status ===
                "Belum Follow Up"
        ).length;


    const followUp =
        prospects.filter(
            item =>
                item.status ===
                "Follow Up"
        ).length;


    const interested =
        prospects.filter(
            item =>
                item.status ===
                "Tertarik"
        ).length;


    if (dataSummaryTotal) {

        dataSummaryTotal.textContent =
            total;

    }


    if (dataSummaryPending) {

        dataSummaryPending.textContent =
            pending;

    }


    if (dataSummaryFollowUp) {

        dataSummaryFollowUp.textContent =
            followUp;

    }


    if (dataSummaryInterested) {

        dataSummaryInterested.textContent =
            interested;

    }

}


/* ============================================================
   RENDER DATA TABLE
============================================================ */

function renderDataTable() {

    if (!dataTable) {
        return;
    }


    const rows =
        getFilteredProspects();


    dataTable.innerHTML =
        "";


    /* --------------------------------------------------------
       EMPTY STATE
    -------------------------------------------------------- */

    if (
        rows.length === 0
    ) {

        if (dataEmptyState) {

            dataEmptyState.classList.remove(
                "hidden"
            );

        }

    } else {

        if (dataEmptyState) {

            dataEmptyState.classList.add(
                "hidden"
            );

        }

    }


    /* --------------------------------------------------------
       ROWS
    -------------------------------------------------------- */

    rows.forEach(
        (
            item,
            index
        ) => {

            const tr =
                document.createElement(
                    "tr"
                );


            tr.className =
                "border-b border-slate-100 hover:bg-slate-50/70 transition";


            tr.innerHTML = `

                <td class="p-3 align-top">
                    <input
                        type="checkbox"
                        class="data-row-checkbox w-4 h-4 rounded border-slate-300 text-cargomii-blue focus:ring-cargomii-blue cursor-pointer"
                        data-id="${escapeHTML(item.id)}"
                        ${selectedDataIds.has(String(item.id)) ? "checked" : ""}
                        aria-label="Pilih ${escapeHTML(item.company || "data prospek")}"
                    >
                </td>

                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-400
                        align-top
                    "
                >
                    ${index + 1}
                </td>


                <td class="p-3 align-top">

                    <p
                        class="
                            text-[10px]
                            font-bold
                            text-cargomii-navy
                            max-w-[220px]
                        "
                    >
                        ${escapeHTML(
                            item.company || "-"
                        )}
                    </p>

                    <p
                        class="
                            text-[8px]
                            text-slate-400
                            mt-1
                        "
                    >
                        ${escapeHTML(
                            item.category ||
                            item.website ||
                            ""
                        )}
                    </p>

                </td>


                <td class="p-3 align-top">

                    <span
                        class="
                            text-[9px]
                            font-bold
                            text-slate-600
                            whitespace-nowrap
                        "
                    >
                        ${escapeHTML(
                            item.phone || "-"
                        )}
                    </span>

                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                    "
                >
                    ${escapeHTML(
                        item.region || "-"
                    )}
                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                    "
                >
                    ${escapeHTML(
                        item.pic || "-"
                    )}
                </td>


                <td
                    class="
                        p-3
                        text-[9px]
                        text-slate-500
                        align-top
                    "
                >
                    ${escapeHTML(
                        item.source || "-"
                    )}
                </td>


                <td class="p-3 align-top">

                    ${getActivityBadges(item)}

                </td>


                <td class="p-3 align-top">

                    ${getStatusBadge(
                        item.status
                    )}

                </td>


                <td class="p-3 align-top">

                    <div
                        class="
                            flex
                            items-center
                            gap-1.5
                        "
                    >

                        <button
                            type="button"
                            data-followup-id="${escapeHTML(item.id)}"
                            title="Follow Up"
                            class="
                                data-followup-btn
                                w-8 h-8
                                rounded-lg
                                bg-blue-50
                                text-cargomii-blue
                                flex
                                items-center
                                justify-center
                                hover:bg-blue-100
                                transition
                            "
                        >
                            <i
                                data-lucide="message-circle"
                                class="w-3.5 h-3.5"
                            ></i>
                        </button>


                        <button
                            type="button"
                            data-delete-id="${escapeHTML(item.id)}"
                            title="Hapus"
                            class="
                                data-delete-btn
                                w-8 h-8
                                rounded-lg
                                bg-red-50
                                text-red-500
                                flex
                                items-center
                                justify-center
                                hover:bg-red-100
                                transition
                            "
                        >
                            <i
                                data-lucide="trash-2"
                                class="w-3.5 h-3.5"
                            ></i>
                        </button>

                    </div>

                </td>

            `;


            dataTable.appendChild(
                tr
            );

        }
    );


    /* --------------------------------------------------------
       ROW CHECKBOXES
    -------------------------------------------------------- */

    dataTable
        .querySelectorAll(".data-row-checkbox")
        .forEach(checkbox => {

            checkbox.addEventListener("change", () => {

                const id = String(checkbox.dataset.id);

                if (checkbox.checked) {
                    selectedDataIds.add(id);
                } else {
                    selectedDataIds.delete(id);
                }

                updateDataBulkActions();

            });

        });


    updateDataBulkActions();


    /* --------------------------------------------------------
       FOLLOW UP BUTTONS
    -------------------------------------------------------- */

    dataTable
        .querySelectorAll(
            ".data-followup-btn"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

                        currentFollowUpId =
                            button.dataset.followupId;


                        openTab(
                            "follow-up"
                        );


                        /*
                           Fungsi ada di Part 3.
                        */

                        if (
                            typeof populateCompanySelect ===
                            "function"
                        ) {

                            populateCompanySelect();

                        }


                        const select =
                            el("companySelect");


                        if (select) {

                            select.value =
                                currentFollowUpId;

                        }


                        if (
                            typeof renderFollowUpCustomer ===
                            "function"
                        ) {

                            renderFollowUpCustomer();

                        }

                    }
                );

            }
        );


    /* --------------------------------------------------------
       DELETE BUTTONS
    -------------------------------------------------------- */

    dataTable
        .querySelectorAll(
            ".data-delete-btn"
        )
        .forEach(
            button => {

                button.addEventListener(
                    "click",
                    () => {

                        openDeleteModal(
                            button.dataset.deleteId
                        );

                    }
                );

            }
        );


    if (window.lucide) {

        lucide.createIcons();

    }

}


/* ============================================================
   SETUP DATA FILTERS
============================================================ */

function setupDataFilters() {

    if (searchInput) {

        searchInput.addEventListener(
            "input",
            renderDataTable
        );

    }


    if (statusFilter) {

        statusFilter.addEventListener(
            "change",
            renderDataTable
        );

    }

}


/* ============================================================
   FIND DUPLICATE PROSPECT

   Prioritas pengecekan:
   1. Nomor HP yang sama setelah dinormalisasi.
   2. Nama perusahaan + wilayah yang sama.
============================================================ */

function normalizeDuplicateText(value) {

    return cleanText(value)
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");

}


function findDuplicateProspect(data, ignoreId = null) {

    const phone = normalizePhone(data?.phone);
    const company = normalizeDuplicateText(data?.company);
    const region = normalizeDuplicateText(data?.region);

    return prospects.find(item => {

        if (
            ignoreId !== null &&
            String(item.id) === String(ignoreId)
        ) {
            return false;
        }

        const samePhone =
            phone &&
            normalizePhone(item.phone) === phone;

        const sameCompanyAndRegion =
            company &&
            region &&
            normalizeDuplicateText(item.company) === company &&
            normalizeDuplicateText(item.region) === region;

        return samePhone || sameCompanyAndRegion;

    }) || null;

}


/* ============================================================
   DATA TABLE BULK SELECTION
============================================================ */

function updateDataBulkActions() {

    const existingIds = new Set(
        prospects.map(item => String(item.id))
    );

    selectedDataIds.forEach(id => {
        if (!existingIds.has(String(id))) {
            selectedDataIds.delete(id);
        }
    });

    const visibleIds = getFilteredProspects()
        .map(item => String(item.id));

    const selectedVisible = visibleIds.filter(
        id => selectedDataIds.has(id)
    ).length;

    if (selectAllDataCheckbox) {
        selectAllDataCheckbox.checked =
            visibleIds.length > 0 &&
            selectedVisible === visibleIds.length;

        selectAllDataCheckbox.indeterminate =
            selectedVisible > 0 &&
            selectedVisible < visibleIds.length;

        selectAllDataCheckbox.disabled =
            visibleIds.length === 0;
    }

    if (deleteSelectedDataBtn) {
        deleteSelectedDataBtn.disabled =
            selectedDataIds.size === 0;
    }

    if (deleteSelectedDataText) {
        deleteSelectedDataText.textContent =
            `Hapus Terpilih (${selectedDataIds.size})`;
    }

}


function setupDataBulkActions() {

    if (selectAllDataCheckbox) {

        selectAllDataCheckbox.addEventListener("change", () => {

            const visibleIds = getFilteredProspects()
                .map(item => String(item.id));

            if (selectAllDataCheckbox.checked) {
                visibleIds.forEach(id => selectedDataIds.add(id));
            } else {
                visibleIds.forEach(id => selectedDataIds.delete(id));
            }


            renderDataTable();

        });

    }

    if (deleteSelectedDataBtn) {

        deleteSelectedDataBtn.addEventListener("click", () => {

            if (selectedDataIds.size === 0) {
                return;
            }

            deleteTargetId = null;
            deleteTargetIds = Array.from(selectedDataIds);

            if (deleteModal) {
                deleteModal.classList.remove("hidden");
            }

        });

    }

}


/* ============================================================
   OPEN DELETE MODAL
============================================================ */

function openDeleteModal(
    id
) {

    const item =
        getProspectById(id);


    if (!item) {
        return;
    }


    deleteTargetId =
        item.id;

    deleteTargetIds = [];


    if (deleteModal) {

        deleteModal.classList.remove(
            "hidden"
        );

    }

}


/* ============================================================
   CLOSE DELETE MODAL
============================================================ */

function closeDeleteModal() {

    deleteTargetId =
        null;

    deleteTargetIds = [];


    if (deleteModal) {

        deleteModal.classList.add(
            "hidden"
        );

    }

}


/* ============================================================
   SETUP DELETE MODAL
============================================================ */

function setupDeleteModal() {

    if (cancelDeleteBtn) {

        cancelDeleteBtn.addEventListener(
            "click",
            closeDeleteModal
        );

    }


    if (confirmDeleteBtn) {

        confirmDeleteBtn.addEventListener(
            "click",
            () => {

                if (!deleteTargetId && deleteTargetIds.length === 0) {

                    closeDeleteModal();

                    return;

                }


                const idsToDelete = deleteTargetIds.length > 0
                    ? deleteTargetIds.map(String)
                    : [String(deleteTargetId)];

                const item = deleteTargetId
                    ? getProspectById(deleteTargetId)
                    : null;

                const deletedCount = idsToDelete.length;


                prospects =
                    prospects.filter(
                        prospect =>
                            !idsToDelete.includes(
                                String(prospect.id)
                            )
                    );


                if (
                    idsToDelete.includes(
                        String(currentFollowUpId)
                    )
                ) {

                    currentFollowUpId =
                        null;

                }

                idsToDelete.forEach(
                    id => selectedDataIds.delete(id)
                );


                saveProspects();

                closeDeleteModal();

                refreshAll();


                showToast(
                    item
                        ? `${item.company} berhasil dihapus.`
                        : `${deletedCount} data berhasil dihapus.`
                );

            }
        );

    }


    /* --------------------------------------------------------
       CLICK BACKDROP
    -------------------------------------------------------- */

    if (deleteModal) {

        deleteModal.addEventListener(
            "click",
            event => {

                if (
                    event.target ===
                    deleteModal
                ) {

                    closeDeleteModal();

                }

            }
        );

    }


    /* --------------------------------------------------------
       ESC
    -------------------------------------------------------- */

    document.addEventListener(
        "keydown",
        event => {

            if (
                event.key === "Escape" &&
                deleteModal &&
                !deleteModal.classList.contains(
                    "hidden"
                )
            ) {

                closeDeleteModal();

            }

        }
    );

}


/* ============================================================
   REFRESH ALL

   Fungsi utama supaya setiap perubahan data langsung
   memperbarui Dashboard + Data Saya.

   Part 3 nanti menambahkan:
   - Follow Up
   - Company select
   - Progress follow up
============================================================ */

function refreshAll() {

    renderMarketingProfile();

    renderDashboard();

    renderDataSummary();

    renderDataTable();


    /* --------------------------------------------------------
       PART 3 FUNCTIONS
    -------------------------------------------------------- */

    if (
        typeof populateCompanySelect ===
        "function"
    ) {

        populateCompanySelect();

    }


    if (
        typeof renderFollowUpProgress ===
        "function"
    ) {

        renderFollowUpProgress();

    }


    if (
        currentFollowUpId &&
        typeof renderFollowUpCustomer ===
        "function"
    ) {

        renderFollowUpCustomer();

    }


    if (window.lucide) {

        lucide.createIcons();

    }

}


/* ============================================================
   BAGIAN 2 SELESAI

   JANGAN TUTUP SCRIPT / JANGAN TAMBAH KODE LAIN.

   APP.JS BAGIAN 3 HARUS LANGSUNG DITEMPEL DI BAWAH INI.

   BAGIAN 3 AKAN BERISI:
   - Follow Up
   - Pilih perusahaan
   - Template pesan
   - WhatsApp
   - SMS
   - Call
   - Tracking aktivitas
   - Status prospek
   - Progress target
   - Export Excel final
   - Profil marketing integration
   - Initialization aplikasi
============================================================ */

/* ============================================================
   CARGOMII MARKETING TOOL
   APP.JS FULL — BAGIAN 3 / 3

   LANJUTAN LANGSUNG DARI BAGIAN 2
============================================================ */


/* ============================================================
   FOLLOW UP ELEMENTS
============================================================ */

const companySelect =
    el("companySelect");

const selectedCustomerInfo =
    el("selectedCustomerInfo");

const selectedCompanyName =
    el("selectedCompanyName");

const selectedCompanyRegion =
    el("selectedCompanyRegion");

const selectedCompanyPic =
    el("selectedCompanyPic");

const customerPhone =
    el("customerPhone");

const copyPhoneBtn =
    el("copyPhoneBtn");

const followUpStatusSelect =
    el("followUpStatusSelect");

const saveFollowUpStatusBtn =
    el("saveFollowUpStatusBtn");

const phoneLinkStatus =
    el("phoneLinkStatus");

const smsStatus =
    el("smsStatus");

const waStatus =
    el("waStatus");

const callStatus =
    el("callStatus");

const templateSelect =
    el("templateSelect");

const messageInput =
    el("messageInput");

const copyBtn =
    el("copyBtn");

const phoneLinkBtn =
    el("phoneLinkBtn");

const whatsappBtn =
    el("whatsappBtn");

const callBtn =
    el("callBtn");

const followUpProgressPercent =
    el("followUpProgressPercent");

const followUpProgressText =
    el("followUpProgressText");

const followUpProgressBar =
    el("followUpProgressBar");

const exportBtn =
    el("exportBtn");


/* ============================================================
   ENSURE ACTIVITIES
============================================================ */

function ensureActivities(item) {

    if (!item) {
        return {
            sms: false,
            whatsapp: false,
            call: false
        };
    }


    if (
        !item.activities ||
        typeof item.activities !== "object"
    ) {

        item.activities = {
            sms: false,
            whatsapp: false,
            call: false
        };

    }


    item.activities.sms =
        Boolean(
            item.activities.sms
        );


    item.activities.whatsapp =
        Boolean(
            item.activities.whatsapp
        );


    item.activities.call =
        Boolean(
            item.activities.call
        );


    return item.activities;

}


/* ============================================================
   POPULATE COMPANY SELECT
============================================================ */

function populateCompanySelect() {

    if (!companySelect) {
        return;
    }


    const oldValue =
        currentFollowUpId ||
        companySelect.value;


    companySelect.innerHTML = `
        <option value="">
            Pilih perusahaan
        </option>
    `;


    prospects.forEach(
        item => {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                item.id;


            option.textContent =
                item.company
                    ? `${item.company} — ${item.phone || "-"}`
                    : item.phone || "Tanpa Nama";


            companySelect.appendChild(
                option
            );

        }
    );


    if (
        oldValue &&
        getProspectById(oldValue)
    ) {

        companySelect.value =
            oldValue;

    }

}


/* ============================================================
   RESET FOLLOW UP CUSTOMER
============================================================ */

function resetFollowUpCustomer() {

    currentFollowUpId =
        null;


    if (selectedCustomerInfo) {

        selectedCustomerInfo.classList.add(
            "hidden"
        );

    }


    if (selectedCompanyName) {

        selectedCompanyName.textContent =
            "-";

    }


    if (selectedCompanyRegion) {

        selectedCompanyRegion.textContent =
            "-";

    }


    if (selectedCompanyPic) {

        selectedCompanyPic.textContent =
            "-";

    }


    if (customerPhone) {

        customerPhone.value =
            "";

    }


    if (followUpStatusSelect) {

        followUpStatusSelect.value =
            "Belum Follow Up";

    }


    if (phoneLinkStatus) {

        phoneLinkStatus.textContent =
            "Belum";

    }


    if (smsStatus) {

        smsStatus.textContent =
            "Belum";

    }


    if (waStatus) {

        waStatus.textContent =
            "Belum";

    }


    if (callStatus) {

        callStatus.textContent =
            "Belum";

    }


    if (messageInput) {

        messageInput.value =
            "";

    }

}


/* ============================================================
   GET SELECTED PROSPECT
============================================================ */

function getSelectedProspect() {

    if (!currentFollowUpId) {

        return null;

    }


    return getProspectById(
        currentFollowUpId
    );

}


/* ============================================================
   RENDER FOLLOW UP CUSTOMER
============================================================ */

function renderFollowUpCustomer() {

    if (!companySelect) {
        return;
    }


    const selectedId =
        companySelect.value ||
        currentFollowUpId;


    if (!selectedId) {

        resetFollowUpCustomer();

        return;

    }


    const item =
        getProspectById(
            selectedId
        );


    if (!item) {

        resetFollowUpCustomer();

        return;

    }


    currentFollowUpId =
        item.id;


    companySelect.value =
        item.id;


    const activities =
        ensureActivities(
            item
        );


    if (selectedCustomerInfo) {

        selectedCustomerInfo.classList.remove(
            "hidden"
        );

    }


    if (selectedCompanyName) {

        selectedCompanyName.textContent =
            item.company || "-";

    }


    if (selectedCompanyRegion) {

        selectedCompanyRegion.textContent =
            item.region ||
            detectRegionFromAddress(
                item.address
            ) ||
            "-";

    }


    if (selectedCompanyPic) {

        selectedCompanyPic.textContent =
            item.pic || "-";

    }


    if (customerPhone) {

        customerPhone.value =
            item.phone || "";

    }


    if (followUpStatusSelect) {

        followUpStatusSelect.value =
            item.status ||
            "Belum Follow Up";

    }


    if (phoneLinkStatus) {

        phoneLinkStatus.textContent =
            item.phone
                ? "Tersedia"
                : "Tidak Ada";

    }


    if (smsStatus) {

        smsStatus.textContent =
            activities.sms
                ? "Sudah"
                : "Belum";

    }


    if (waStatus) {

        waStatus.textContent =
            activities.whatsapp
                ? "Sudah"
                : "Belum";

    }


    if (callStatus) {

        callStatus.textContent =
            activities.call
                ? "Sudah"
                : "Belum";

    }


    updateMessageTemplate();

}


/* ============================================================
   MARK ACTIVITY
============================================================ */

function markProspectActivity(
    prospectId,
    activity
) {

    const item =
        getProspectById(
            prospectId
        );


    if (!item) {
        return;
    }


    ensureActivities(
        item
    );


    if (
        activity === "sms" ||
        activity === "whatsapp" ||
        activity === "call"
    ) {

        item.activities[
            activity
        ] = true;

    }


    /*
       Begitu marketing mulai menghubungi customer,
       status otomatis menjadi Follow Up.

       Status Tertarik / Tidak Tertarik tidak dioverwrite.
    */

    if (
        !item.status ||
        item.status ===
            "Belum Follow Up"
    ) {

        item.status =
            "Follow Up";

    }


    item.updatedAt =
        new Date().toISOString();


    saveProspects();

    refreshAll();

}


/* ============================================================
   MARKETING NAME FOR TEMPLATE
============================================================ */

function getMarketingName() {

    return (
        cleanText(
            marketingProfile.name
        ) ||
        "Tim Cargomii"
    );

}


/* ============================================================
   MESSAGE TEMPLATE
============================================================ */

function generateMessage(templateName) {
    if (templateName === "custom") return "";

    const marketingName = cleanText(marketingProfile.name) || "Nama Marketing";
    const marketingPhone = normalizePhone(marketingProfile.phone) || "Nomor Marketing";
    return `Cargomii via Kapal Cepat. Promo min 1Kg Sudah Bisa Kirim dan Diantar Sampai Alamat
Tujuan :
Jkt- Sorong Manokwari Jayapura
tarif Rp 13.500/kg
${marketingPhone}
${marketingName}
Cargomii

Salam,
${marketingName}
Cargomii`;
}


/* ============================================================
   UPDATE MESSAGE TEMPLATE
============================================================ */

function updateMessageTemplate() {

    if (!messageInput) {
        return;
    }


    const template =
        templateSelect?.value ||
        "kapal-cepat";


    /*
       CUSTOM tidak dioverwrite.
    */

    if (
        template === "custom"
    ) {

        return;

    }


    messageInput.value =
        generateMessage(
            template
        );

}


/* ============================================================
   COPY TEXT HELPER
============================================================ */

async function copyText(
    text
) {

    if (!text) {
        return false;
    }


    try {

        if (
            navigator.clipboard &&
            window.isSecureContext
        ) {

            await navigator.clipboard.writeText(
                text
            );

            return true;

        }


        const textarea =
            document.createElement(
                "textarea"
            );


        textarea.value =
            text;


        textarea.style.position =
            "fixed";


        textarea.style.opacity =
            "0";


        document.body.appendChild(
            textarea
        );


        textarea.focus();

        textarea.select();


        const success =
            document.execCommand(
                "copy"
            );


        textarea.remove();


        return success;

    } catch (error) {

        console.error(
            "COPY ERROR:",
            error
        );


        return false;

    }

}


/* ============================================================
   SETUP FOLLOW UP
============================================================ */

function setupFollowUp() {

    /* --------------------------------------------------------
       COMPANY
    -------------------------------------------------------- */

    if (companySelect) {

        companySelect.addEventListener(
            "change",
            () => {

                currentFollowUpId =
                    companySelect.value ||
                    null;


                renderFollowUpCustomer();

            }
        );

    }


    /* --------------------------------------------------------
       TEMPLATE
    -------------------------------------------------------- */

    if (templateSelect) {

        templateSelect.addEventListener(
            "change",
            () => {

                if (
                    templateSelect.value ===
                    "custom"
                ) {

                    if (messageInput) {

                        messageInput.value =
                            "";

                        messageInput.focus();

                    }

                    return;

                }


                updateMessageTemplate();

            }
        );

    }


    /* --------------------------------------------------------
       COPY MESSAGE
    -------------------------------------------------------- */

    if (copyBtn) {

        copyBtn.addEventListener(
            "click",
            async () => {

                const message =
                    cleanText(
                        messageInput?.value
                    );


                if (!message) {

                    showToast(
                        "Belum ada pesan untuk disalin.",
                        "error"
                    );

                    return;

                }


                const success =
                    await copyText(
                        messageInput.value
                    );


                if (success) {

                    showToast(
                        "Pesan berhasil disalin."
                    );

                } else {

                    showToast(
                        "Pesan gagal disalin.",
                        "error"
                    );

                }

            }
        );

    }


    /* --------------------------------------------------------
       COPY PHONE
    -------------------------------------------------------- */

    if (copyPhoneBtn) {

        copyPhoneBtn.addEventListener(
            "click",
            async () => {

                const item =
                    getSelectedProspect();


                if (
                    !item ||
                    !item.phone
                ) {

                    showToast(
                        "Nomor customer belum tersedia.",
                        "error"
                    );

                    return;

                }


                const success =
                    await copyText(
                        item.phone
                    );


                if (success) {

                    showToast(
                        "Nomor berhasil disalin."
                    );

                } else {

                    showToast(
                        "Nomor gagal disalin.",
                        "error"
                    );

                }

            }
        );

    }


    /* --------------------------------------------------------
       WHATSAPP
    -------------------------------------------------------- */

    if (whatsappBtn) {

        whatsappBtn.addEventListener(
            "click",
            () => {

                const item =
                    getSelectedProspect();


                if (!item) {

                    showToast(
                        "Pilih perusahaan terlebih dahulu.",
                        "error"
                    );

                    return;

                }


                const phone =
                    toWhatsAppNumber(
                        item.phone
                    );


                if (!phone) {

                    showToast(
                        "Nomor WhatsApp tidak tersedia.",
                        "error"
                    );

                    return;

                }


                const message =
                    messageInput?.value ||
                    "";


                /*
                   Buka tab dahulu agar popup blocker
                   tidak terlalu agresif.
                */

                const url =
                    `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;


                markProspectActivity(
                    item.id,
                    "whatsapp"
                );


                window.open(
                    url,
                    "_blank",
                    "noopener,noreferrer"
                );

            }
        );

    }


    /* --------------------------------------------------------
       SMS
    -------------------------------------------------------- */

    if (phoneLinkBtn) {

        phoneLinkBtn.addEventListener(
            "click",
            () => {

                const item =
                    getSelectedProspect();


                if (!item) {

                    showToast(
                        "Pilih perusahaan terlebih dahulu.",
                        "error"
                    );

                    return;

                }


                const phone =
                    normalizePhone(
                        item.phone
                    );


                if (!phone) {

                    showToast(
                        "Nomor customer tidak tersedia.",
                        "error"
                    );

                    return;

                }


                const message =
                    messageInput?.value ||
                    "";


                markProspectActivity(
                    item.id,
                    "sms"
                );


                /*
                   Di desktop ini tergantung aplikasi
                   Phone Link / SMS handler Windows.
                */

                window.location.href =
                    `sms:${phone}?body=${encodeURIComponent(message)}`;

            }
        );

    }


    /* --------------------------------------------------------
       CALL
    -------------------------------------------------------- */

    if (callBtn) {

        callBtn.addEventListener(
            "click",
            () => {

                const item =
                    getSelectedProspect();


                if (!item) {

                    showToast(
                        "Pilih perusahaan terlebih dahulu.",
                        "error"
                    );

                    return;

                }


                const phone =
                    normalizePhone(
                        item.phone
                    );


                if (!phone) {

                    showToast(
                        "Nomor customer tidak tersedia.",
                        "error"
                    );

                    return;

                }


                markProspectActivity(
                    item.id,
                    "call"
                );


                window.location.href =
                    `tel:${phone}`;

            }
        );

    }


    /* --------------------------------------------------------
       SAVE STATUS
    -------------------------------------------------------- */

    if (saveFollowUpStatusBtn) {

        saveFollowUpStatusBtn.addEventListener(
            "click",
            () => {

                const item =
                    getSelectedProspect();


                if (!item) {

                    showToast(
                        "Pilih perusahaan terlebih dahulu.",
                        "error"
                    );

                    return;

                }


                const status =
                    followUpStatusSelect?.value ||
                    "Belum Follow Up";


                const allowedStatuses = [
                    "Belum Follow Up",
                    "Follow Up",
                    "Tertarik",
                    "Tidak Tertarik"
                ];


                item.status =
                    allowedStatuses.includes(
                        status
                    )
                        ? status
                        : "Belum Follow Up";


                item.updatedAt =
                    new Date().toISOString();


                saveProspects();

                refreshAll();


                showToast(
                    "Status prospek berhasil diperbarui."
                );

            }
        );

    }

}


/* ============================================================
   FOLLOW UP PROGRESS
============================================================ */

function renderFollowUpProgress() {

    const target =
        Math.max(
            Number(
                marketingProfile.target || 0
            ),
            0
        );


    const total =
        prospects.length;


    const percentage =
        target > 0
            ? Math.min(
                Math.round(
                    (total / target) * 100
                ),
                100
            )
            : 0;


    if (followUpProgressPercent) {

        followUpProgressPercent.textContent =
            `${percentage}%`;

    }


    if (followUpProgressText) {

        followUpProgressText.textContent =
            `${total} / ${target} Data`;

    }


    if (followUpProgressBar) {

        followUpProgressBar.style.width =
            `${percentage}%`;

    }

}


/* ============================================================
   EXPORT HELPERS
============================================================ */

function getExportDate(
    item
) {

    const source =
        item.createdAt ||
        item.updatedAt;


    if (!source) {

        return new Date();

    }


    const date =
        new Date(
            source
        );


    if (
        Number.isNaN(
            date.getTime()
        )
    ) {

        return new Date();

    }


    return date;

}


/* ============================================================
   EXPORT STATUS / KETERANGAN
============================================================ */

function getExportStatus(
    item
) {

    switch (
        item.status
    ) {

        case "Tertarik":
            return "Tertarik";


        case "Tidak Tertarik":
            return "Tidak Tertarik";


        case "Follow Up":
            return "Follow Up";


        default:
            return "Belum Follow Up";

    }

}


/* ============================================================
   BORDER HELPER
============================================================ */

function excelBorder() {

    return {

        top: {
            style: "thin",
            color: {
                argb: "FFB8C0CC"
            }
        },

        left: {
            style: "thin",
            color: {
                argb: "FFB8C0CC"
            }
        },

        bottom: {
            style: "thin",
            color: {
                argb: "FFB8C0CC"
            }
        },

        right: {
            style: "thin",
            color: {
                argb: "FFB8C0CC"
            }
        }

    };

}


/* ============================================================
   EXPORT EXCEL

   HASIL:

   GOOGLE BUSINESS NAMA MARKETING

   NO
   NAMA
   TANGGAL
   NOMER WHATSAPP
   FOLLOW UP
      SMS
      WA
      CALL
   KET
============================================================ */

async function exportProspectsToExcel() {

    if (
        !Array.isArray(prospects) ||
        prospects.length === 0
    ) {

        showToast(
            "Belum ada data untuk diexport.",
            "error"
        );

        return;

    }


    if (
        typeof ExcelJS ===
        "undefined"
    ) {

        console.error(
            "ExcelJS is not defined."
        );


        showToast(
            "ExcelJS belum dimuat di index.html.",
            "error"
        );

        return;

    }


    try {

        /* ====================================================
           WORKBOOK
        ==================================================== */

        const workbook =
            new ExcelJS.Workbook();


        workbook.creator =
            "Cargomii Marketing Tool";


        workbook.company =
            "Cargomii";


        workbook.created =
            new Date();


        /* ====================================================
           WORKSHEET
        ==================================================== */

        const worksheet =
            workbook.addWorksheet(
                "GOOGLE BUSINESS",
                {
                    views: [
                        {
                            state: "frozen",
                            ySplit: 3
                        }
                    ]
                }
            );


        /* ====================================================
           PAGE SETUP
        ==================================================== */

        worksheet.pageSetup = {

            orientation:
                "landscape",

            paperSize:
                9,

            fitToPage:
                true,

            fitToWidth:
                1,

            fitToHeight:
                0,

            margins: {

                left: 0.25,
                right: 0.25,
                top: 0.5,
                bottom: 0.5,
                header: 0.2,
                footer: 0.2

            }

        };


        /* ====================================================
           COLUMN WIDTH
        ==================================================== */

        worksheet.columns = [

            {
                key: "no",
                width: 7
            },

            {
                key: "nama",
                width: 38
            },

            {
                key: "tanggal",
                width: 16
            },

            {
                key: "whatsapp",
                width: 22
            },

            {
                key: "sms",
                width: 10
            },

            {
                key: "wa",
                width: 10
            },

            {
                key: "call",
                width: 10
            },

            {
                key: "ket",
                width: 25
            }

        ];


        /* ====================================================
           MARKETING NAME
        ==================================================== */

        const marketingName =
            (
                marketingProfile.name ||
                "MARKETING"
            )
                .trim()
                .toUpperCase();


        /* ====================================================
           TITLE
        ==================================================== */

        worksheet.mergeCells(
            "A1:H1"
        );


        const titleCell =
            worksheet.getCell(
                "A1"
            );


        titleCell.value =
            `GOOGLE BUSINESS ${marketingName}`;


        titleCell.font = {

            name:
                "Arial",

            size:
                16,

            bold:
                true,

            color: {
                argb:
                    "FFFFFFFF"
            }

        };


        titleCell.alignment = {

            horizontal:
                "center",

            vertical:
                "middle"

        };


        titleCell.fill = {

            type:
                "pattern",

            pattern:
                "solid",

            fgColor: {
                argb:
                    "FF063B91"
            }

        };


        titleCell.border =
            excelBorder();


        worksheet.getRow(
            1
        ).height = 30;


        /* ====================================================
           MERGED HEADER
        ==================================================== */

        worksheet.mergeCells(
            "A2:A3"
        );

        worksheet.mergeCells(
            "B2:B3"
        );

        worksheet.mergeCells(
            "C2:C3"
        );

        worksheet.mergeCells(
            "D2:D3"
        );

        worksheet.mergeCells(
            "E2:G2"
        );

        worksheet.mergeCells(
            "H2:H3"
        );


        /* ====================================================
           HEADER VALUES
        ==================================================== */

        worksheet.getCell(
            "A2"
        ).value = "NO";


        worksheet.getCell(
            "B2"
        ).value = "NAMA";


        worksheet.getCell(
            "C2"
        ).value = "TANGGAL";


        worksheet.getCell(
            "D2"
        ).value = "NOMER WHATSAPP";


        worksheet.getCell(
            "E2"
        ).value = "FOLLOW UP";


        worksheet.getCell(
            "E3"
        ).value = "SMS";


        worksheet.getCell(
            "F3"
        ).value = "WA";


        worksheet.getCell(
            "G3"
        ).value = "CALL";


        worksheet.getCell(
            "H2"
        ).value = "KET";


        /* ====================================================
           HEADER STYLE
        ==================================================== */

        for (
            let rowNumber = 2;
            rowNumber <= 3;
            rowNumber++
        ) {

            const row =
                worksheet.getRow(
                    rowNumber
                );


            row.height =
                25;


            for (
                let columnNumber = 1;
                columnNumber <= 8;
                columnNumber++
            ) {

                const cell =
                    row.getCell(
                        columnNumber
                    );


                cell.font = {

                    name:
                        "Arial",

                    size:
                        10,

                    bold:
                        true,

                    color: {
                        argb:
                            "FFFFFFFF"
                    }

                };


                cell.fill = {

                    type:
                        "pattern",

                    pattern:
                        "solid",

                    fgColor: {
                        argb:
                            "FF071A3D"
                    }

                };


                cell.alignment = {

                    horizontal:
                        "center",

                    vertical:
                        "middle",

                    wrapText:
                        true

                };


                cell.border =
                    excelBorder();

            }

        }


        /* ====================================================
           DATA
        ==================================================== */

        prospects.forEach(
            (
                item,
                index
            ) => {

                const activities =
                    ensureActivities(
                        item
                    );


                const rowNumber =
                    index + 4;


                const row =
                    worksheet.getRow(
                        rowNumber
                    );


                /* ============================================
                   NO
                ============================================ */

                row.getCell(
                    1
                ).value =
                    index + 1;


                /* ============================================
                   NAMA
                ============================================ */

                row.getCell(
                    2
                ).value =
                    item.company ||
                    "-";


                /* ============================================
                   TANGGAL
                ============================================ */

                row.getCell(
                    3
                ).value =
                    getExportDate(
                        item
                    );


                row.getCell(
                    3
                ).numFmt =
                    "dd/mm/yyyy";


                /* ============================================
                   NOMOR WA

                   Dipaksa STRING supaya:
                   081234...
                   tidak berubah menjadi:
                   81234...
                ============================================ */

                row.getCell(
                    4
                ).value =
                    String(
                        item.phone ||
                        ""
                    );


                row.getCell(
                    4
                ).numFmt =
                    "@";


                /* ============================================
                   SMS
                ============================================ */

                row.getCell(
                    5
                ).value =
                    activities.sms
                        ? "✓"
                        : "";


                /* ============================================
                   WA
                ============================================ */

                row.getCell(
                    6
                ).value =
                    activities.whatsapp
                        ? "✓"
                        : "";


                /* ============================================
                   CALL
                ============================================ */

                row.getCell(
                    7
                ).value =
                    activities.call
                        ? "✓"
                        : "";


                /* ============================================
                   KET
                ============================================ */

                row.getCell(
                    8
                ).value =
                    getExportStatus(
                        item
                    );


                row.height =
                    23;


                /* ============================================
                   GENERAL STYLE
                ============================================ */

                for (
                    let columnNumber = 1;
                    columnNumber <= 8;
                    columnNumber++
                ) {

                    const cell =
                        row.getCell(
                            columnNumber
                        );


                    cell.font = {

                        name:
                            "Arial",

                        size:
                            10,

                        color: {
                            argb:
                                "FF1E293B"
                        }

                    };


                    cell.alignment = {

                        vertical:
                            "middle",

                        horizontal:
                            (
                                columnNumber === 2 ||
                                columnNumber === 8
                            )
                                ? "left"
                                : "center",

                        wrapText:
                            true

                    };


                    cell.border =
                        excelBorder();


                    /*
                       Alternating row.
                    */

                    if (
                        index % 2 !== 0
                    ) {

                        cell.fill = {

                            type:
                                "pattern",

                            pattern:
                                "solid",

                            fgColor: {
                                argb:
                                    "FFF8FAFC"
                            }

                        };

                    }

                }


                /* ============================================
                   CHECKMARK STYLE
                ============================================ */

                [
                    5,
                    6,
                    7
                ].forEach(
                    columnNumber => {

                        const cell =
                            row.getCell(
                                columnNumber
                            );


                        if (
                            cell.value === "✓"
                        ) {

                            cell.font = {

                                name:
                                    "Arial",

                                size:
                                    12,

                                bold:
                                    true,

                                color: {
                                    argb:
                                        "FF15803D"
                                }

                            };

                        }

                    }
                );


                /* ============================================
                   STATUS STYLE
                ============================================ */

                const statusCell =
                    row.getCell(
                        8
                    );


                if (
                    item.status ===
                    "Tertarik"
                ) {

                    statusCell.font = {

                        name:
                            "Arial",

                        size:
                            10,

                        bold:
                            true,

                        color: {
                            argb:
                                "FF15803D"
                        }

                    };

                }


                else if (
                    item.status ===
                    "Tidak Tertarik"
                ) {

                    statusCell.font = {

                        name:
                            "Arial",

                        size:
                            10,

                        bold:
                            true,

                        color: {
                            argb:
                                "FFDC2626"
                        }

                    };

                }


                else if (
                    item.status ===
                    "Follow Up"
                ) {

                    statusCell.font = {

                        name:
                            "Arial",

                        size:
                            10,

                        bold:
                            true,

                        color: {
                            argb:
                                "FFD97706"
                        }

                    };

                }

            }
        );


        /* ====================================================
           AUTO FILTER

           Header kompleks E:G membuat filter lebih aman
           dimulai pada row 3.
        ==================================================== */

        worksheet.autoFilter = {
            from: "A3",
            to: "H3"
        };


        /* ====================================================
           PRINT AREA
        ==================================================== */

        const lastRow =
            prospects.length + 3;


        worksheet.pageSetup.printArea =
            `A1:H${lastRow}`;


        /* ====================================================
           FOOTER
        ==================================================== */

        worksheet.headerFooter = {

            oddFooter:
                `&LCargomii Marketing Tool&C${marketingName}&RPage &P of &N`

        };


        /* ====================================================
           CREATE BUFFER
        ==================================================== */

        const buffer =
            await workbook.xlsx.writeBuffer();


        /* ====================================================
           FILE NAME
        ==================================================== */

        const safeMarketingName =
            (
                marketingProfile.name ||
                "Marketing"
            )
                .trim()
                .replace(
                    /[^a-z0-9]+/gi,
                    "_"
                )
                .replace(
                    /^_+|_+$/g,
                    ""
                ) ||
                "Marketing";


        const now =
            new Date();


        const day =
            String(
                now.getDate()
            ).padStart(
                2,
                "0"
            );


        const month =
            String(
                now.getMonth() + 1
            ).padStart(
                2,
                "0"
            );


        const year =
            now.getFullYear();


        const fileName =
            `GOOGLE_BUSINESS_${safeMarketingName}_${day}-${month}-${year}.xlsx`;


        /* ====================================================
           BLOB
        ==================================================== */

        const blob =
            new Blob(
                [buffer],
                {
                    type:
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
            );


        /* ====================================================
           DOWNLOAD

           FileSaver kalau tersedia.
           Kalau tidak, native browser.
        ==================================================== */

        if (
            typeof saveAs ===
            "function"
        ) {

            saveAs(
                blob,
                fileName
            );

        } else {

            const url =
                URL.createObjectURL(
                    blob
                );


            const anchor =
                document.createElement(
                    "a"
                );


            anchor.href =
                url;


            anchor.download =
                fileName;


            anchor.style.display =
                "none";


            document.body.appendChild(
                anchor
            );


            anchor.click();


            anchor.remove();


            setTimeout(
                () => {

                    URL.revokeObjectURL(
                        url
                    );

                },
                1000
            );

        }


        showToast(
            "Excel berhasil dibuat."
        );

    } catch (error) {

        console.error(
            "EXPORT EXCEL ERROR:",
            error
        );


        showToast(
            "Gagal membuat file Excel. Cek Console browser.",
            "error"
        );

    }

}


/* ============================================================
   SETUP EXPORT
============================================================ */

/* ============================================================
   UNDUH KONTAK (VCF / vCard 3.0)
============================================================ */

function escapeVCardText(value) {
    return String(value ?? "")
        .replace(/\r\n|\r|\n/g, "\n")
        .replace(/\\/g, "\\\\")
        .replace(/;/g, "\\;")
        .replace(/,/g, "\\,")
        .replace(/\n/g, "\\n");
}

function exportContactsToVcf() {
    const selected = selectedDataIds.size > 0;
    const rows = selected
        ? prospects.filter(item => selectedDataIds.has(String(item.id)))
        : getFilteredProspects();
    const seenPhones = new Set();
    const cards = [];

    for (const item of rows) {
        const phone = normalizePhone(item.phone);
        if (!isValidMobilePhone(phone) || seenPhones.has(phone)) continue;
        seenPhones.add(phone);

        const company = cleanText(item.company);
        const pic = cleanText(item.pic);
        const name = [pic, company].filter(Boolean).join(" - ") || phone;
        const lines = [
            "BEGIN:VCARD",
            "VERSION:3.0",
            `FN:${escapeVCardText(name)}`,
            `N:${escapeVCardText(name)};;;;`,
            `TEL;TYPE=CELL:+62${phone.slice(1)}`
        ];
        if (company) lines.push(`ORG:${escapeVCardText(company)}`);
        if (item.region) lines.push(`ADR;TYPE=WORK:;;;${escapeVCardText(item.region)};;;`);
        if (item.notes) lines.push(`NOTE:${escapeVCardText(item.notes)}`);
        lines.push("END:VCARD");
        cards.push(lines.join("\r\n"));
    }

    if (!cards.length) {
        showToast("Tidak ada nomor HP valid untuk diunduh.", "error");
        return;
    }

    const blob = new Blob([cards.join("\r\n") + "\r\n"], {
        type: "text/vcard;charset=utf-8"
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `Cargomii_Kontak_${new Date().toISOString().slice(0, 10)}.vcf`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    showToast(`${cards.length} kontak berhasil diunduh sebagai VCF.`);
}

function setupExport() {

    el("exportVcfBtn")?.addEventListener("click", exportContactsToVcf);

    if (!exportBtn) {
        return;
    }


    exportBtn.addEventListener(
        "click",
        exportProspectsToExcel
    );

}


/* ============================================================
   NORMALIZE ALL EXISTING DATA
============================================================ */

function normalizeExistingData() {

    let changed =
        false;


    prospects =
        prospects.map(
            item => {

                const before =
                    JSON.stringify(
                        item
                    );


                const normalized =
                    normalizeProspect(
                        item
                    );


                ensureActivities(
                    normalized
                );


                const after =
                    JSON.stringify(
                        normalized
                    );


                if (
                    before !== after
                ) {

                    changed =
                        true;

                }


                return normalized;

            }
        );


    if (changed) {

        saveProspects();

    }

}


/* ============================================================
   INITIAL TAB
============================================================ */

function getInitialTab() {

    /*
       Kalau profile marketing belum dibuat,
       buka dashboard saja.

       Tidak dipaksa ke halaman lain agar UI tetap sederhana.
    */

    const activeButton =
        document.querySelector(
            ".nav-btn.active[data-tab]"
        );


    if (
        activeButton?.dataset?.tab
    ) {

        return activeButton.dataset.tab;

    }


    return "dashboard";

}


/* ============================================================
   APP INITIALIZATION
============================================================ */

function initializeApp() {

    try {

        /* ----------------------------------------------------
           1. LOAD DATA
        ---------------------------------------------------- */

        loadStorage();


        /* ----------------------------------------------------
           2. NORMALIZE DATA
        ---------------------------------------------------- */

        normalizeExistingData();


        /* ----------------------------------------------------
           3. SETUP NAVIGATION
        ---------------------------------------------------- */

        setupNavigation();


        /* ----------------------------------------------------
           4. PROFILE
        ---------------------------------------------------- */

        setupMarketingProfile();


        /* ----------------------------------------------------
           5. PHONE INPUT
        ---------------------------------------------------- */

        setupPhoneInputs();


        /* ----------------------------------------------------
           6. MANUAL FORM
        ---------------------------------------------------- */

        setupManualForm();


        /* ----------------------------------------------------
           7. EXCEL
        ---------------------------------------------------- */

        setupExcelUpload();


        /* ----------------------------------------------------
           8. DATA FILTER
        ---------------------------------------------------- */

        setupDataFilters();

        setupDataBulkActions();


        /* ----------------------------------------------------
           9. DELETE MODAL
        ---------------------------------------------------- */

        setupDeleteModal();


        /* ----------------------------------------------------
           10. FOLLOW UP
        ---------------------------------------------------- */

        setupFollowUp();


        /* ----------------------------------------------------
           11. EXPORT
        ---------------------------------------------------- */

        setupExport();


        /* ----------------------------------------------------
           12. INITIAL RENDER
        ---------------------------------------------------- */

        refreshAll();


        /* ----------------------------------------------------
           13. INITIAL TAB
        ---------------------------------------------------- */

        openTab(
            getInitialTab()
        );


        /* ----------------------------------------------------
           14. ICON
        ---------------------------------------------------- */

        if (window.lucide) {

            lucide.createIcons();

        }


        console.log(
            "Cargomii Marketing Tool berhasil dimuat."
        );

    } catch (error) {

        console.error(
            "INITIALIZATION ERROR:",
            error
        );


        showToast(
            "Terjadi error saat menjalankan aplikasi. Buka Console untuk melihat detail.",
            "error"
        );

    }

}


/* ============================================================
   RUN APP

   Karena app.js idealnya dipasang sebelum </body>,
   pengecekan ini membuatnya tetap aman jika script
   dipasang di <head>.
============================================================ */

if (
    document.readyState ===
    "loading"
) {

    document.addEventListener(
        "DOMContentLoaded",
        initializeApp
    );

} else {

    initializeApp();

}


/* ============================================================
   END
   CARGOMII MARKETING TOOL
   APP.JS FULL 1 + 2 + 3
============================================================ */
