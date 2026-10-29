import { readFile } from 'node:fs/promises';

import { Actor, log } from 'apify';
import { PlaywrightCrawler } from 'crawlee';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

await Actor.init();

log.info('Actor initialized successfully.');

let supabase;


// --------------------------------------------------
// VALIDATION
// --------------------------------------------------

function validateDate(value, fieldName) {
    const datePattern =
        /^(0?[1-9]|1[0-2])\/(0?[1-9]|[12]\d|3[01])\/\d{4}$/;

    if (!datePattern.test(value)) {
        throw new Error(
            `${fieldName} must use MM/DD/YYYY format. `
            + `Received: ${value}`,
        );
    }
}


function validateTime(value, fieldName) {
    const timePattern =
        /^(0?[1-9]|1[0-2]):[0-5]\d$/;

    if (!timePattern.test(value)) {
        throw new Error(
            `${fieldName} must use HH:MM 12-hour format. `
            + `Received: ${value}`,
        );
    }
}


// --------------------------------------------------
// DATE HELPERS
// --------------------------------------------------

const CENTRAL_TIME_ZONE = 'America/Chicago';


function getCentralDateParts(date) {
    const parts = new Intl.DateTimeFormat(
        'en-US',
        {
            timeZone: CENTRAL_TIME_ZONE,
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
        },
    ).formatToParts(date);

    const lookup = Object.fromEntries(
        parts
            .filter((part) => part.type !== 'literal')
            .map((part) => [part.type, Number(part.value)]),
    );

    return {
        year: lookup.year,
        month: lookup.month,
        day: lookup.day,
    };
}


function formatDateParts({ year, month, day }) {
    return (
        `${String(month).padStart(2, '0')}/`
        + `${String(day).padStart(2, '0')}/`
        + `${year}`
    );
}


function calculateStartDate(today = new Date()) {
    const { year, month, day } =
        getCentralDateParts(today);

    const calendarDate = new Date(
        Date.UTC(year, month - 1, day),
    );

    calendarDate.setUTCDate(
        calendarDate.getUTCDate() - 1,
    );

    return formatDateParts({
        year: calendarDate.getUTCFullYear(),
        month: calendarDate.getUTCMonth() + 1,
        day: calendarDate.getUTCDate(),
    });
}


function calculateEndDate(today = new Date()) {
    return calculateStartDate(today);
}


function formatBusinessDate(value) {
    const [month, day, year] = value.split('/');

    return (
        `${year}-`
        + `${month.padStart(2, '0')}-`
        + `${day.padStart(2, '0')}`
    );
}


// --------------------------------------------------
// NUMBER CONVERSION
// --------------------------------------------------

function toNumber(value) {
    if (
        value === null
        || value === undefined
        || value === ''
    ) {
        return null;
    }

    const cleaned = String(value)
        .replace(/[$,\s]/g, '')
        .trim();

    if (!cleaned) {
        return null;
    }

    const number = Number(cleaned);

    return Number.isFinite(number)
        ? number
        : null;
}


// --------------------------------------------------
// DISCOUNT REASON PARSER
// --------------------------------------------------

function isDiscountSectionHeader(row) {
    const label =
        String(row?.[0] ?? '').trim();

    const columnHeader =
        String(row?.[1] ?? '').trim();

    if (!label || !columnHeader) {
        return false;
    }

    // The next report section starts with a label
    // such as Qty, Amount, or Hours instead of a number.
    return toNumber(columnHeader) === null;
}


function extractDiscountReasons(rows) {

    const startIndex = rows.findIndex((row) => {
        const firstColumn =
            String(row?.[0] ?? '')
                .trim()
                .toUpperCase();

        return firstColumn === 'DISCOUNT REASON';
    });

    if (startIndex === -1) {
        throw new Error(
            'Unable to find DISCOUNT REASON in Operations CSV.',
        );
    }

    log.info(
        `DISCOUNT REASON found at CSV row ${startIndex + 1}.`,
    );

    const discountReasons = [];

    for (
        let index = startIndex + 1;
        index < rows.length;
        index++
    ) {

        const row = rows[index];

        const reason =
            String(row?.[0] ?? '').trim();

        if (isDiscountSectionHeader(row)) {
            log.info(
                'Reached the next Operations section. '
                + 'Discount Reason extraction complete.',
            );

            break;
        }

        if (!reason) {
            continue;
        }

        discountReasons.push({
            reason,
            qty: toNumber(row?.[1]),
            total: toNumber(row?.[2]),
        });
    }

    log.info(
        `Extracted ${discountReasons.length} `
        + 'discount reason row(s).',
    );

    return discountReasons;
}


function buildDiscountReasonDictionary(discountReasons) {
    const discountReasonsByName = {};

    for (const entry of discountReasons) {
        discountReasonsByName[entry.reason] = {
            qty: entry.qty,
            total: entry.total,
        };
    }

    return discountReasonsByName;
}


// --------------------------------------------------
// ESTABLISHMENT MAPPING
// --------------------------------------------------

const ESTABLISHMENT_MAP = {
    Lampasas: '41 | Lampasas',
    Leander: '42 | Leander',
    'Marble Falls': '29 | Marble Falls',
    Temple: '62 | Temple',
};


// --------------------------------------------------
// MAIN
// --------------------------------------------------

try {

    log.info('Starting Revel Operations Discount Reason actor.');

    // --------------------------------------------------
    // SUPABASE
    // --------------------------------------------------

    log.info('Reading Supabase configuration.');

    const supabaseUrl =
        process.env.SUPABASE_URL
        || 'https://ongqhvokcwceqgnetonq.supabase.co';

    const supabaseServiceRoleKey =
        process.env.SUPABASE_SERVICE_ROLE_KEY;

    log.info('Supabase environment check.', {
        hasUrl: Boolean(supabaseUrl),
        hasServiceRoleKey: Boolean(supabaseServiceRoleKey),
    });

    if (!supabaseServiceRoleKey) {
        throw new Error(
            'SUPABASE_SERVICE_ROLE_KEY is not configured.',
        );
    }

    supabase = createClient(
        supabaseUrl,
        supabaseServiceRoleKey,
        {
            auth: {
                persistSession: false,
                autoRefreshToken: false,
            },
        },
    );

    log.info('Supabase client initialized.');

    // --------------------------------------------------
    // ACTOR INPUT
    // --------------------------------------------------

    log.info('Reading Actor input.');

    const input = await Actor.getInput();

    log.info('Actor input received.', {
        hasInput: Boolean(input),
        inputKeys: input ? Object.keys(input) : [],
    });

    const {
        url =
            'https://laynes.revelup.com/reports/operations/',

        username,
        password,

        establishment = 'Leander',

        override_flag = false,
        override_startDate,
        override_endDate,

        startTime,
        startMeridiem,

        endTime,
        endMeridiem,

    } = input ?? {};


    // --------------------------------------------------
    // INPUT VALIDATION
    // --------------------------------------------------

    let startDate;
    let endDate;

    if (override_flag) {
        if (!override_startDate || !override_endDate) {
            throw new Error(
                'override_startDate and override_endDate '
                + 'are required when override_flag is true.',
            );
        }

        startDate = override_startDate;
        endDate = override_endDate;
    } else {
        const today = new Date();

        startDate = calculateStartDate(today);
        endDate = calculateEndDate(today);
    }

    log.info(
        override_flag
            ? 'Using override dates.'
            : 'Calculated dates for US Central Time. '
                + 'Start and end date are both yesterday.',
        {
            override_flag,
            startDate,
            endDate,
        },
    );

    if (!username || !password) {
        throw new Error(
            'Both username and password are required.',
        );
    }

    if (
        !startTime
        || !startMeridiem
        || !endTime
        || !endMeridiem
    ) {
        throw new Error(
            'Start and end times and AM/PM values '
            + 'are required.',
        );
    }

    validateDate(startDate, 'startDate');
    validateDate(endDate, 'endDate');

    validateTime(startTime, 'startTime');
    validateTime(endTime, 'endTime');


    const normalizedStartMeridiem =
        startMeridiem.trim().toUpperCase();

    const normalizedEndMeridiem =
        endMeridiem.trim().toUpperCase();


    if (
        !['AM', 'PM'].includes(normalizedStartMeridiem)
        || !['AM', 'PM'].includes(normalizedEndMeridiem)
    ) {
        throw new Error(
            'startMeridiem and endMeridiem '
            + 'must be AM or PM.',
        );
    }


    const targetEstablishment =
        establishment.trim();


    const targetEstablishmentTreeText =
        ESTABLISHMENT_MAP[targetEstablishment];


    if (!targetEstablishmentTreeText) {
        throw new Error(
            `Unsupported establishment: ${targetEstablishment}`,
        );
    }


    log.info(
        `Target establishment: ${targetEstablishment}`,
    );


    log.info(
        `Target establishment tree entry: `
        + `${targetEstablishmentTreeText}`,
    );

    log.info(
        `Requested Operations range: `
        + `${startDate} ${startTime} `
        + `${normalizedStartMeridiem} through `
        + `${endDate} ${endTime} `
        + `${normalizedEndMeridiem}`,
    );


    // --------------------------------------------------
    // CRAWLER
    // --------------------------------------------------

    const crawler = new PlaywrightCrawler({

        maxRequestsPerCrawl: 1,
        maxRequestRetries: 0,
        requestHandlerTimeoutSecs: 240,


        async requestHandler({ page, request }) {

            // ==================================================
            // OPEN REVEL
            // ==================================================

            log.info(
                `Opening Revel portal: ${request.url}`,
            );

            await page.goto(
                request.url,
                {
                    waitUntil: 'domcontentloaded',
                    timeout: 30_000,
                },
            );


            // ==================================================
            // LOGIN - USERNAME
            // ==================================================

            const usernameField =
                page.locator('#username');

            await usernameField.waitFor({
                state: 'visible',
                timeout: 15_000,
            });

            await usernameField.fill(username);

            log.info(
                'Username entered. Clicking Continue.',
            );

            await page
                .getByRole(
                    'button',
                    {
                        name: 'Continue',
                        exact: true,
                    },
                )
                .click();


            // ==================================================
            // LOGIN - PASSWORD
            // ==================================================

            const passwordField =
                page.locator(
                    'input[type="password"]',
                );

            await passwordField.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            await passwordField.fill(password);


            const loginButton =
                page
                    .locator(
                        'button[type="submit"]:visible, '
                        + 'input[type="submit"]:visible',
                    )
                    .last();


            await loginButton.waitFor({
                state: 'visible',
                timeout: 15_000,
            });


            log.info(
                'Password entered. Logging into Revel.',
            );

            await loginButton.click();


            await passwordField.waitFor({
                state: 'hidden',
                timeout: 30_000,
            });


            await page.waitForLoadState(
                'domcontentloaded',
            );


            log.info(
                `Login completed. Current URL: `
                + `${page.url()}`,
            );


            // ==================================================
            // NAVIGATE TO OPERATIONS
            // ==================================================

            if (
                !page.url().includes(
                    '/reports/operations',
                )
            ) {

                log.info(
                    `Navigating to Operations report: `
                    + `${url}`,
                );

                await page.goto(
                    url,
                    {
                        waitUntil:
                            'domcontentloaded',

                        timeout: 30_000,
                    },
                );
            }


            // ==================================================
            // ESTABLISHMENT
            // ==================================================

            const establishmentText =
                page.locator(
                    '[data-cy="header-establishment-text"]',
                );


            await establishmentText.waitFor({
                state: 'visible',
                timeout: 20_000,
            });


            let currentEstablishment =
                (
                    await establishmentText
                        .textContent()
                )?.trim()
                || 'Unknown';


            log.info(
                `Current establishment: `
                + `${currentEstablishment}`,
            );


            // ==================================================
            // SELECT ESTABLISHMENT IF NECESSARY
            // ==================================================

            if (
                currentEstablishment
                !== targetEstablishment
            ) {

                await establishmentText.click();

                log.info(
                    'Establishment panel opened.',
                );


                // ----------------------------------------------
                // Sort by establishment number
                // ----------------------------------------------

                const sortByEstablishmentNumber =
                    page.locator(
                        '.btn.by-id',
                    );


                if (
                    await sortByEstablishmentNumber
                        .count()
                    > 0
                ) {

                    await sortByEstablishmentNumber
                        .first()
                        .click();

                    log.info(
                        'Sorted establishment tree '
                        + 'by establishment number.',
                    );
                }


                // ----------------------------------------------
                // Expand all folders
                // ----------------------------------------------

                const expandAll =
                    page.locator(
                        'span.expand-all',
                    );


                if (
                    await expandAll.count()
                    > 0
                ) {

                    await expandAll
                        .first()
                        .click();

                    log.info(
                        'Expanded all establishment '
                        + 'folders.',
                    );
                }


                await page.waitForTimeout(1000);


                // ----------------------------------------------
                // Find requested establishment
                // ----------------------------------------------

                const escapedTreeText =
                    targetEstablishmentTreeText
                        .replace(
                            /[.*+?^${}()|[\]\\]/g,
                            '\\$&',
                        );


                const targetOption =
                    page
                        .locator(
                            'span.fancytree-title',
                        )
                        .filter({
                            hasText:
                                new RegExp(
                                    `^\\s*${escapedTreeText}\\s*$`,
                                ),
                        })
                        .first();


                await targetOption.waitFor({
                    state: 'visible',
                    timeout: 30_000,
                });


                const selectedTreeText =
                    (
                        await targetOption
                            .textContent()
                    )?.trim();


                log.info(
                    `Found establishment tree entry: `
                    + `${selectedTreeText}`,
                );


                log.info(
                    `Selecting establishment: `
                    + `${targetEstablishment}`,
                );


                // FancyTree titles can be visible without a simple
                // Playwright click triggering Revel's selection handler.
                // Dispatch a native mouse sequence against the exact title.
                await targetOption.scrollIntoViewIfNeeded();

                const clickedTreeText =
                    await targetOption.evaluate(
                        (element) => {
                            const text =
                                element.textContent?.trim()
                                || '';

                            element.dispatchEvent(
                                new MouseEvent(
                                    'mousedown',
                                    {
                                        bubbles: true,
                                        cancelable: true,
                                        view: window,
                                    },
                                ),
                            );

                            element.dispatchEvent(
                                new MouseEvent(
                                    'mouseup',
                                    {
                                        bubbles: true,
                                        cancelable: true,
                                        view: window,
                                    },
                                ),
                            );

                            element.dispatchEvent(
                                new MouseEvent(
                                    'click',
                                    {
                                        bubbles: true,
                                        cancelable: true,
                                        view: window,
                                    },
                                ),
                            );

                            return text;
                        },
                    );


                log.info(
                    `Dispatched native click sequence on `
                    + `establishment tree entry: `
                    + `${clickedTreeText}`,
                );


                const headerImmediatelyAfterClick =
                    (
                        await establishmentText
                            .textContent()
                    )?.trim();


                log.info(
                    `Establishment header immediately `
                    + `after click: `
                    + `${headerImmediatelyAfterClick}`,
                );


                await page.waitForTimeout(2000);


                const headerAfterTwoSeconds =
                    (
                        await establishmentText
                            .textContent()
                    )?.trim();


                log.info(
                    `Establishment header after 2 seconds: `
                    + `${headerAfterTwoSeconds}`,
                );


                log.info(
                    `Current URL after establishment click: `
                    + `${page.url()}`,
                );
            }


            // ==================================================
            // VERIFY ESTABLISHMENT
            // ==================================================

            const establishmentVerificationDeadline =
                Date.now() + 30_000;


            let verifiedHeaderText = 'Unknown';


            while (
                Date.now()
                < establishmentVerificationDeadline
            ) {

                verifiedHeaderText =
                    (
                        await establishmentText
                            .textContent()
                    )?.trim()
                    || 'Unknown';


                if (
                    verifiedHeaderText
                    === targetEstablishment
                ) {
                    break;
                }


                await page.waitForTimeout(250);
            }


            if (
                verifiedHeaderText
                !== targetEstablishment
            ) {

                const panelVisible =
                    await page
                        .locator(
                            'span.fancytree-title:visible',
                        )
                        .count();


                throw new Error(
                    `Establishment selection did not complete. `
                    + `Clicked "${targetEstablishmentTreeText}", `
                    + `but header is "${verifiedHeaderText}". `
                    + `Visible FancyTree titles: ${panelVisible}. `
                    + `URL: ${page.url()}.`,
                );
            }


            log.info(
                `Establishment header verification passed: `
                + `${verifiedHeaderText}`,
            );


            currentEstablishment =
                (
                    await establishmentText
                        .textContent()
                )?.trim();


            if (
                currentEstablishment
                !== targetEstablishment
            ) {

                throw new Error(
                    `Expected establishment `
                    + `"${targetEstablishment}", `
                    + `but found `
                    + `"${currentEstablishment}".`,
                );
            }


            log.info(
                `Verified establishment: `
                + `${currentEstablishment}`,
            );


            // ==================================================
            // DATE RANGE PICKER
            // ==================================================

            const dateRangeDropdown =
                page.locator(
                    '.report-date-row '
                    + '.ico-f-to-down',
                );


            await dateRangeDropdown.waitFor({
                state: 'visible',
                timeout: 20_000,
            });


            log.info(
                'Opening Operations date-range picker.',
            );


            await dateRangeDropdown.click();


            const visibleDatePicker =
                page.locator(
                    '.daterangepicker:visible',
                );


            await visibleDatePicker.waitFor({
                state: 'visible',
                timeout: 20_000,
            });


            // ==================================================
            // SET DATE RANGE
            // ==================================================

            const pickerResult =
                await page.evaluate(
                    ({
                        startDateValue,
                        startTimeValue,
                        startMeridiemValue,
                        endDateValue,
                        endTimeValue,
                        endMeridiemValue,
                    }) => {

                        const $ =
                            window.jQuery;

                        const moment =
                            window.moment;


                        if (!$) {
                            throw new Error(
                                'jQuery is not '
                                + 'available.',
                            );
                        }


                        if (!moment) {
                            throw new Error(
                                'Moment.js is not '
                                + 'available.',
                            );
                        }


                        const candidates =
                            $('*').filter(
                                function findPicker() {
                                    return Boolean(
                                        $(this).data(
                                            'daterangepicker',
                                        ),
                                    );
                                },
                            );


                        if (
                            candidates.length === 0
                        ) {
                            throw new Error(
                                'Unable to locate '
                                + 'Revel date picker.',
                            );
                        }


                        let picker = null;


                        candidates.each(
                            function selectPicker() {

                                const candidate =
                                    $(this).data(
                                        'daterangepicker',
                                    );


                                if (
                                    !picker
                                    && candidate
                                        ?.container
                                    && candidate
                                        .container
                                        .is(':visible')
                                ) {
                                    picker =
                                        candidate;
                                }
                            },
                        );


                        if (!picker) {
                            picker =
                                $(candidates[0])
                                    .data(
                                        'daterangepicker',
                                    );
                        }


                        const startDateTime =
                            moment(
                                `${startDateValue} `
                                + `${startTimeValue} `
                                + `${startMeridiemValue}`,

                                'MM/DD/YYYY '
                                + 'hh:mm A',

                                true,
                            );


                        const endDateTime =
                            moment(
                                `${endDateValue} `
                                + `${endTimeValue} `
                                + `${endMeridiemValue}`,

                                'MM/DD/YYYY '
                                + 'hh:mm A',

                                true,
                            );


                        if (
                            !startDateTime.isValid()
                        ) {
                            throw new Error(
                                'Invalid start '
                                + 'date/time.',
                            );
                        }


                        if (
                            !endDateTime.isValid()
                        ) {
                            throw new Error(
                                'Invalid end '
                                + 'date/time.',
                            );
                        }


                        if (
                            endDateTime.isBefore(
                                startDateTime,
                            )
                        ) {
                            throw new Error(
                                'End date/time cannot '
                                + 'be before start.',
                            );
                        }


                        picker.setStartDate(
                            startDateTime,
                        );

                        picker.setEndDate(
                            endDateTime,
                        );


                        if (
                            typeof picker.updateView
                            === 'function'
                        ) {
                            picker.updateView();
                        }


                        if (
                            typeof picker
                                .updateCalendars
                            === 'function'
                        ) {
                            picker
                                .updateCalendars();
                        }


                        if (
                            typeof picker
                                .updateFormInputs
                            === 'function'
                        ) {
                            picker
                                .updateFormInputs();
                        }


                        return {
                            startDate:
                                picker.startDate
                                    .format(
                                        'MM/DD/YYYY '
                                        + 'hh:mm A',
                                    ),

                            endDate:
                                picker.endDate
                                    .format(
                                        'MM/DD/YYYY '
                                        + 'hh:mm A',
                                    ),
                        };
                    },

                    {
                        startDateValue:
                            startDate,

                        startTimeValue:
                            startTime,

                        startMeridiemValue:
                            normalizedStartMeridiem,

                        endDateValue:
                            endDate,

                        endTimeValue:
                            endTime,

                        endMeridiemValue:
                            normalizedEndMeridiem,
                    },
                );


            log.info(
                `Picker range: `
                + `${pickerResult.startDate} `
                + `through `
                + `${pickerResult.endDate}`,
            );


            // ==================================================
            // APPLY DATE RANGE
            // ==================================================

            await page.evaluate(() => {

                const $ =
                    window.jQuery;


                const candidates =
                    $('*').filter(
                        function findPicker() {
                            return Boolean(
                                $(this).data(
                                    'daterangepicker',
                                ),
                            );
                        },
                    );


                let picker = null;


                candidates.each(
                    function selectPicker() {

                        const candidate =
                            $(this).data(
                                'daterangepicker',
                            );


                        if (
                            !picker
                            && candidate
                                ?.container
                            && candidate
                                .container
                                .is(':visible')
                        ) {
                            picker =
                                candidate;
                        }
                    },
                );


                if (
                    !picker
                    && candidates.length > 0
                ) {
                    picker =
                        $(candidates[0])
                            .data(
                                'daterangepicker',
                            );
                }


                if (!picker) {
                    throw new Error(
                        'Unable to locate '
                        + 'daterangepicker '
                        + 'during Apply.',
                    );
                }


                if (
                    typeof picker.clickApply
                    !== 'function'
                ) {
                    throw new Error(
                        'Date picker does not '
                        + 'expose clickApply().',
                    );
                }


                picker.clickApply();
            });


            await visibleDatePicker.waitFor({
                state: 'hidden',
                timeout: 30_000,
            });


            log.info(
                'Date range applied. '
                + 'Waiting for Operations report.',
            );


            // ==================================================
            // WAIT FOR DATE DISPLAY
            // ==================================================

            await page.waitForFunction(
                ({
                    expectedStartDate,
                    expectedEndDate,
                }) => {

                    const normalizeDate =
                        (value) => {

                            const match =
                                String(value).match(
                                    /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
                                );

                            if (!match) {
                                return null;
                            }

                            const [
                                ,
                                month,
                                day,
                                year,
                            ] = match;

                            return (
                                `${month.padStart(2, '0')}/`
                                + `${day.padStart(2, '0')}/`
                                + `${year}`
                            );
                        };


                    const reportDateRow =
                        document.querySelector(
                            '.report-date-row',
                        );


                    if (!reportDateRow) {
                        return false;
                    }


                    const displayedDates =
                        (
                            reportDateRow
                                .textContent
                            ?? ''
                        )
                            .match(
                                /\d{1,2}\/\d{1,2}\/\d{4}/g,
                            )
                            ?.map(
                                normalizeDate,
                            );


                    if (
                        !displayedDates
                        || displayedDates.length < 2
                    ) {
                        return false;
                    }


                    return (
                        displayedDates[0]
                        === normalizeDate(
                            expectedStartDate,
                        )

                        &&

                        displayedDates[1]
                        === normalizeDate(
                            expectedEndDate,
                        )
                    );
                },

                {
                    expectedStartDate:
                        startDate,

                    expectedEndDate:
                        endDate,
                },

                {
                    timeout: 90_000,
                    polling: 500,
                },
            );


            // ==================================================
            // WAIT FOR LOADERS
            // ==================================================

            await page.waitForFunction(
                () => {

                    const loadingElements =
                        document
                            .querySelectorAll(
                                [
                                    '.loading',
                                    '.loader',
                                    '.spinner',
                                    '.loading-mask',
                                    '.blockUI',
                                    '.fa-spinner',
                                    '.icon-spinner',
                                    '[class*='
                                    + '"loading-indicator"]',
                                ].join(','),
                            );


                    return [
                        ...loadingElements,
                    ].every(
                        (element) => {

                            const style =
                                window
                                    .getComputedStyle(
                                        element,
                                    );

                            const bounds =
                                element
                                    .getBoundingClientRect();


                            return (
                                style.display
                                    === 'none'

                                || style.visibility
                                    === 'hidden'

                                || style.opacity
                                    === '0'

                                || bounds.width
                                    === 0

                                || bounds.height
                                    === 0
                            );
                        },
                    );
                },

                undefined,

                {
                    timeout: 90_000,
                    polling: 500,
                },
            );


            await page.waitForTimeout(
                1500,
            );


            log.info(
                'Operations report refreshed.',
            );


            // ==================================================
            // OPEN EXPORT MENU AND FIND VISIBLE CSV
            // ==================================================

            const exportMenuButton =
                page.locator(
                    '.header-more .button-more, '
                    + '.button.button-square.button-more',
                ).first();


            await exportMenuButton.waitFor({
                state: 'visible',
                timeout: 30_000,
            });


            log.info(
                'Opening Operations Export menu.',
            );


            await exportMenuButton.click();


            const visibleCsvExportLink =
                page
                    .locator(
                        'a[href^="data.csv?"]:visible',
                    )
                    .filter({
                        hasText:
                            /^\s*CSV\s*$/,
                    })
                    .first();


            await visibleCsvExportLink.waitFor({
                state: 'visible',
                timeout: 30_000,
            });


            const csvHref =
                await visibleCsvExportLink
                    .getAttribute('href');


            if (!csvHref) {
                throw new Error(
                    'Visible Operations CSV export '
                    + 'does not contain an href.',
                );
            }


            log.info(
                `Visible Operations CSV export found: `
                + `${csvHref}`,
            );


            // ==================================================
            // DOWNLOAD CSV THROUGH VISIBLE EXPORT CONTROL
            // ==================================================

            const downloadPromise =
                page.waitForEvent(
                    'download',
                    {
                        timeout: 60_000,
                    },
                );


            await visibleCsvExportLink.click();


            const download =
                await downloadPromise;


            const downloadFailure =
                await download.failure();


            if (downloadFailure) {
                throw new Error(
                    `CSV download failed: `
                    + `${downloadFailure}`,
                );
            }


            const temporaryFilePath =
                await download.path();


            if (!temporaryFilePath) {
                throw new Error(
                    'Playwright did not provide '
                    + 'a temporary CSV path.',
                );
            }


            const csvBuffer =
                await readFile(
                    temporaryFilePath,
                );


            log.info(
                `Downloaded Operations CSV: `
                + `${csvBuffer.length} bytes.`,
            );


            // ==================================================
            // PARSE CSV
            // ==================================================

            const workbook =
                XLSX.read(
                    csvBuffer,
                    {
                        type: 'buffer',
                        raw: true,
                    },
                );


            const firstSheetName =
                workbook.SheetNames[0];


            if (!firstSheetName) {
                throw new Error(
                    'Operations CSV contains '
                    + 'no worksheet.',
                );
            }


            const worksheet =
                workbook.Sheets[
                    firstSheetName
                ];


            const rows =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: null,
                        raw: true,
                    },
                );


            log.info(
                `Parsed ${rows.length} `
                + `CSV row(s).`,
            );


            // ==================================================
            // DISCOUNT REASONS
            // ==================================================

            const discountReasons =
                extractDiscountReasons(rows);


            log.info(
                'Discount Reasons extracted.',
                {
                    count: discountReasons.length,
                    reasons: discountReasons,
                },
            );


            // ==================================================
            // BUSINESS DATE
            // ==================================================

            const formattedBusinessDate =
                formatBusinessDate(
                    startDate,
                );


            // ==================================================
            // DISCOUNT REASON ROW
            // ==================================================

            const discountReasonDictionary =
                buildDiscountReasonDictionary(
                    discountReasons,
                );

            const discountReasonRow = {

                id:
                    `${formattedBusinessDate}_`
                    + `${currentEstablishment}`,

                location:
                    currentEstablishment,

                business_date:
                    formattedBusinessDate,

                discount_reasons:
                    discountReasonDictionary,

                extracted_at:
                    new Date().toISOString(),
            };


            const {
                data: savedDiscountRows,
                error: discountSupabaseError,
            } =
                await supabase
                    .from(
                        'daily-revel-discount-reason',
                    )
                    .upsert(
                        discountReasonRow,
                        {
                            onConflict: 'id',
                        },
                    )
                    .select();


            if (discountSupabaseError) {
                throw new Error(
                    'Unable to write Discount Reasons '
                    + 'to Supabase: '
                    + `${discountSupabaseError.message}`,
                );
            }


            log.info(
                `Discount Reasons saved to Supabase: `
                + `${discountReasonRow.id}`,
                discountReasonDictionary,
            );


            // ==================================================
            // ACTOR DATASET LOG
            // ==================================================

            await Actor.pushData({

                status: 'success',

                report:
                    'Operations - Discount Reasons',

                location:
                    currentEstablishment,

                businessDate:
                    formattedBusinessDate,

                discountReasonRecordId:
                    discountReasonRow.id,

                discountReasonCount:
                    Object.keys(
                        discountReasonDictionary,
                    ).length,

                savedDiscountRows:
                    savedDiscountRows?.length ?? 0,

                timestamp:
                    new Date()
                        .toISOString(),

            });
        },


        // ======================================================
        // FAILURE HANDLER
        // ======================================================

        async failedRequestHandler(
            {
                page,
                request,
            },
            error,
        ) {

            log.error(
                `Revel Operations extraction `
                + `failed: ${error.message}`,
            );


            await Actor.pushData({

                status: 'failed',

                report:
                    'Operations - Discount Reasons',

                portalUrl:
                    page
                        ? page.url()
                        : request.url,

                location:
                    targetEstablishment,

                startDate,
                endDate,

                timestamp:
                    new Date()
                        .toISOString(),

                message:
                    error.message,
            });
        },
    });


    // --------------------------------------------------
    // RUN
    // --------------------------------------------------

    log.info(`Starting crawler with URL: ${url}`);

    await crawler.run([
        url,
    ]);

    log.info('Crawler finished.');

} finally {

    await Actor.exit();

}