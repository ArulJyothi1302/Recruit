import 'dotenv/config';
import { ApifyClient } from 'apify-client';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import crypto from 'crypto';

// ============================================================
// CONFIGURATION
// ============================================================

const client = new ApifyClient({
    token: process.env.APIFY_API_TOKEN,
});

const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
    },
});

// ============================================================
// JOB SEARCH CONFIG
// ============================================================

// Only these 4 portals.
// DO NOT add Google Jobs or other expensive actors.
const PORTALS = {
    LinkedIn: 'valig/linkedin-jobs-scraper',
    Naukri: 'epicscrapers/naukri-scraper',
    Hirist: 'hipersoft/hirist-scraper',
    Indeed: 'automation-lab/indeed-scraper',
};

// Your target locations
const TARGET_LOCATIONS = [
    'Tamil Nadu',
    'Chennai',
    'Bengaluru',
    'Hyderabad',
    'Coimbatore',
];

// Search broad enough to find different naming variations.
const SEARCH_KEYWORDS =
    'React Node.js Full Stack TypeScript JavaScript AWS';

// Since cron runs every 2 days, 3 days gives some safety overlap.
const SEARCH_DAYS = 3;

// Experience
const EXPERIENCE_MIN = 2;
const EXPERIENCE_MAX = 6;

// Maximum jobs returned per scraper execution.
const MAX_ITEMS = 50;

// ============================================================
// STACK MATCHING
// ============================================================

// Strong technologies from your target stack
const STRONG_SKILLS = [
    'react',
    'react.js',
    'reactjs',
    'node',
    'node.js',
    'nodejs',
    'typescript',
    'javascript',
    'next.js',
    'nextjs',
    'express',
    'express.js',
    'aws',
];

// Supporting technologies
const SUPPORTING_SKILLS = [
    'mongodb',
    'postgresql',
    'postgres',
    'mysql',
    'redis',
    'rest api',
    'restful',
    'graphql',
    'docker',
    'git',
    'github',
];

// Job titles that are highly relevant
const GOOD_TITLE_TERMS = [
    'full stack',
    'fullstack',
    'software engineer',
    'software developer',
    'react developer',
    'reactjs developer',
    'node developer',
    'nodejs developer',
    'frontend engineer',
    'front end engineer',
    'backend engineer',
    'back end engineer',
    'web developer',
    'web engineer',
];

// Technologies/roles you specifically don't want
const EXCLUDED_TERMS = [
    'angular',
    'angularjs',
    'java developer',
    'java engineer',
    'spring boot',
    'spring framework',
    'hibernate',
    '.net developer',
    '.net engineer',
    'dotnet developer',
    'dot net developer',
    'php developer',
    'php engineer',
    'python developer',
    'python engineer',
    'django developer',
    'flask developer',
    'ios developer',
    'android developer',
    'flutter developer',
    'react native developer',
];

// ============================================================
// HELPERS
// ============================================================

function generateMD5(value) {
    return crypto
        .createHash('md5')
        .update(String(value))
        .digest('hex');
}

function safeDate(value) {
    if (!value) {
        return null;
    }

    const date = new Date(value);

    if (Number.isNaN(date.getTime())) {
        return null;
    }

    return date;
}

function normalizeText(value) {
    return String(value || '')
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();
}

function containsTerm(text, term) {
    return normalizeText(text).includes(normalizeText(term));
}

// ============================================================
// JOB SCORING
// ============================================================

function analyzeJob(job) {
    const title = normalizeText(job.Job_Title);
    const description = normalizeText(job.Description_Text);

    const combinedText = `${title} ${description}`;

    let score = 0;

    const matchedStrongSkills = [];
    const matchedSupportingSkills = [];
    const matchedGoodTitles = [];
    const matchedExcludedTerms = [];

    // --------------------------------------------------------
    // Check excluded technologies first
    // --------------------------------------------------------

    for (const term of EXCLUDED_TERMS) {
        if (containsTerm(combinedText, term)) {
            matchedExcludedTerms.push(term);
        }
    }

    // If unwanted stack is found, heavily penalize.
    if (matchedExcludedTerms.length > 0) {
        score -= 100;
    }

    // --------------------------------------------------------
    // Strong skills
    // --------------------------------------------------------

    for (const skill of STRONG_SKILLS) {
        if (containsTerm(combinedText, skill)) {
            matchedStrongSkills.push(skill);
            score += 10;
        }
    }

    // --------------------------------------------------------
    // Supporting skills
    // --------------------------------------------------------

    for (const skill of SUPPORTING_SKILLS) {
        if (containsTerm(combinedText, skill)) {
            matchedSupportingSkills.push(skill);
            score += 3;
        }
    }

    // --------------------------------------------------------
    // Good job titles
    // --------------------------------------------------------

    for (const term of GOOD_TITLE_TERMS) {
        if (containsTerm(title, term)) {
            matchedGoodTitles.push(term);
            score += 15;
        }
    }

    // --------------------------------------------------------
    // Full-stack bonus
    // --------------------------------------------------------

    const hasReact = /react|react\.js|reactjs/i.test(combinedText);
    const hasNode = /node|node\.js|nodejs/i.test(combinedText);
    const hasTypeScript = /typescript/i.test(combinedText);
    const hasAWS = /\baws\b|amazon web services/i.test(combinedText);

    if (hasReact && hasNode) {
        score += 30;
    }

    if (hasReact && hasNode && hasTypeScript) {
        score += 20;
    }

    if (hasReact && hasNode && hasAWS) {
        score += 20;
    }

    // --------------------------------------------------------
    // Determine priority
    // --------------------------------------------------------

    let priority = 'Low';

    if (score >= 80) {
        priority = 'Very High';
    } else if (score >= 55) {
        priority = 'High';
    } else if (score >= 30) {
        priority = 'Medium';
    }

    return {
        score,
        priority,
        matchedStrongSkills,
        matchedSupportingSkills,
        matchedGoodTitles,
        matchedExcludedTerms,
    };
}

// ============================================================
// APIFY DATASET HELPER
// ============================================================

async function getDatasetItems(run) {
    if (!run || !run.defaultDatasetId) {
        return [];
    }

    const dataset = client.dataset(run.defaultDatasetId);

    const result = await dataset.listItems({
        limit: MAX_ITEMS,
    });

    return result.items || [];
}

// ============================================================
// LINKEDIN
// ============================================================

async function scrapeLinkedIn(location) {
    console.log(`🔵 LinkedIn → ${location}`);

    try {
        const run = await client
            .actor(PORTALS.LinkedIn)
            .call({
                keywords: SEARCH_KEYWORDS,
                location,
                datePosted: 'r259200',
                limit: MAX_ITEMS,
            });

        const items = await getDatasetItems(run);

        return {
            portal: 'LinkedIn',
            location,
            items,
        };
    } catch (error) {
        console.error(
            `❌ LinkedIn failed for ${location}:`,
            error.message
        );

        return {
            portal: 'LinkedIn',
            location,
            items: [],
        };
    }
}

// ============================================================
// NAUKRI
// ============================================================

async function scrapeNaukri(location) {
    console.log(`🟢 Naukri → ${location}`);

    try {
        const run = await client
            .actor(PORTALS.Naukri)
            .call({
                keyword: SEARCH_KEYWORDS,
                location,
                experienceMin: EXPERIENCE_MIN,
                experienceMax: EXPERIENCE_MAX,
                postedWithinDays: SEARCH_DAYS,
            });

        const items = await getDatasetItems(run);

        return {
            portal: 'Naukri',
            location,
            items,
        };
    } catch (error) {
        console.error(
            `❌ Naukri failed for ${location}:`,
            error.message
        );

        return {
            portal: 'Naukri',
            location,
            items: [],
        };
    }
}

// ============================================================
// HIRIST
// ============================================================

async function scrapeHirist(location) {
    console.log(`🟠 Hirist → ${location}`);

    try {
        const run = await client
            .actor(PORTALS.Hirist)
            .call({
                queries: [SEARCH_KEYWORDS],
                location,
                maxDaysOld: SEARCH_DAYS,
            });

        const items = await getDatasetItems(run);

        return {
            portal: 'Hirist',
            location,
            items,
        };
    } catch (error) {
        console.error(
            `❌ Hirist failed for ${location}:`,
            error.message
        );

        return {
            portal: 'Hirist',
            location,
            items: [],
        };
    }
}

// ============================================================
// INDEED
// ============================================================

async function scrapeIndeed(location) {
    console.log(`🟣 Indeed → ${location}`);

    try {
        const run = await client
            .actor(PORTALS.Indeed)
            .call({
                query: SEARCH_KEYWORDS,
                location,
                country: 'IN',
                maxDaysOld: SEARCH_DAYS,
                maxItems: MAX_ITEMS,
            });

        const items = await getDatasetItems(run);

        return {
            portal: 'Indeed',
            location,
            items,
        };
    } catch (error) {
        console.error(
            `❌ Indeed failed for ${location}:`,
            error.message
        );

        return {
            portal: 'Indeed',
            location,
            items: [],
        };
    }
}

// ============================================================
// NORMALIZE DIFFERENT PORTAL FORMATS
// ============================================================

function normalizeJob(job, portal, searchLocation) {
    let title = '';
    let company = '';
    let date = null;
    let link = '';
    let description = '';
    let location = '';

    if (portal === 'LinkedIn') {
        title = job.title;
        company = job.companyName;
        date = job.postedAt;
        link = job.url;
        description = job.description;
        location = job.location;
    }

    if (portal === 'Naukri') {
        title = job.title;
        company = job.companyName;
        date = job.postedDate;
        link = job.jobUrl;
        description = job.jobDescription;
        location = job.location;
    }

    if (portal === 'Indeed') {
        title = job.jobTitle;
        company = job.company;
        date = job.date;
        link = job.url;
        description = job.description;
        location = job.location;
    }

    if (portal === 'Hirist') {
        title = job.title;
        company = job.company;
        date = job.date;
        link = job.url;
        description = job.description;
        location = job.location;
    }

    if (!link) {
        return null;
    }

    const parsedDate = safeDate(date);

    return {
        Job_Title: title || 'N/A',
        Company: company || 'N/A',
        Source: portal,
        Date_Posted: parsedDate || new Date(),
        Location: location || searchLocation,
        Experience: `${EXPERIENCE_MIN}-${EXPERIENCE_MAX} years`,
        Stack_Match: '',
        Match_Score: 0,
        Priority: 'Low',
        Direct_Apply_Link: link,
        Description_Text: description || '',
    };
}

// ============================================================
// MAIN SCRAPER
// ============================================================

async function aggregateAndFormatJobs() {
    console.log('');
    console.log('======================================================');
    console.log('🚀 DEEP L2 JOB SEARCH STARTED');
    console.log('======================================================');
    console.log('');

    console.log(`🔎 Search: ${SEARCH_KEYWORDS}`);
    console.log(`📍 Locations: ${TARGET_LOCATIONS.join(', ')}`);
    console.log(`📅 Search window: ${SEARCH_DAYS} days`);
    console.log(`🧰 Portals: LinkedIn, Naukri, Hirist, Indeed`);
    console.log('');

    const allJobs = [];

    // ========================================================
    // STATE
    // ========================================================

    const store = await client
        .keyValueStores()
        .getOrCreate('l2-job-scraper-state-v3');

    const kvClient = client.keyValueStore(store.id);

    const seenRecord = await kvClient.getRecord(
        'SEEN_JOB_HASHES'
    );

    const seenHashes = new Set(
        seenRecord?.value || []
    );

    // ========================================================
    // SCRAPE
    // ========================================================

    let successfulRuns = 0;

    for (const location of TARGET_LOCATIONS) {
        console.log('');
        console.log(`📍 Processing location: ${location}`);
        console.log('---------------------------------------------');

        // Run sequentially instead of launching everything
        // simultaneously. This reduces rate-limit pressure.
        const results = [];

        results.push(
            await scrapeLinkedIn(location)
        );

        results.push(
            await scrapeNaukri(location)
        );

        results.push(
            await scrapeHirist(location)
        );

        results.push(
            await scrapeIndeed(location)
        );

        for (const result of results) {
            successfulRuns++;

            for (const rawJob of result.items) {
                const job = normalizeJob(
                    rawJob,
                    result.portal,
                    result.location
                );

                if (job) {
                    allJobs.push(job);
                }
            }

            console.log(
                `   ${result.portal}: ${result.items.length} jobs`
            );
        }
    }

    console.log('');
    console.log(
        `📦 Raw jobs collected: ${allJobs.length}`
    );

    // ========================================================
    // CUTOFF DATE
    // ========================================================

    const cutoffDate = new Date();

    cutoffDate.setDate(
        cutoffDate.getDate() - SEARCH_DAYS
    );

    // ========================================================
    // DEDUPLICATE WITHIN CURRENT RUN
    // ========================================================

    const currentRunHashes = new Set();

    // ========================================================
    // FILTER + SCORE
    // ========================================================

    const matchedJobs = [];

    for (const job of allJobs) {
        // ----------------------------------------------------
        // Date filter
        // ----------------------------------------------------

        const isRecent =
            job.Date_Posted >= cutoffDate;

        if (!isRecent) {
            continue;
        }

        // ----------------------------------------------------
        // Analyze
        // ----------------------------------------------------

        const analysis = analyzeJob(job);

        // ----------------------------------------------------
        // We need React + Node OR Full Stack + strong skill
        // ----------------------------------------------------

        const titleAndDescription =
            `${job.Job_Title} ${job.Description_Text}`;

        const hasReact =
            /react|react\.js|reactjs/i.test(
                titleAndDescription
            );

        const hasNode =
            /node|node\.js|nodejs/i.test(
                titleAndDescription
            );

        const hasFullStack =
            /full\s*stack|fullstack/i.test(
                titleAndDescription
            );

        const hasRelevantStrongSkill =
            analysis.matchedStrongSkills.length >= 2;

        const relevant =
            (hasReact && hasNode) ||
            (hasFullStack && hasRelevantStrongSkill);

        if (!relevant) {
            continue;
        }

        // ----------------------------------------------------
        // Excluded technology check
        // ----------------------------------------------------

        if (
            analysis.matchedExcludedTerms.length > 0
        ) {
            continue;
        }

        // ----------------------------------------------------
        // Minimum score
        // ----------------------------------------------------

        if (analysis.score < 30) {
            continue;
        }

        // ----------------------------------------------------
        // URL deduplication
        // ----------------------------------------------------

        const urlHash = generateMD5(
            job.Direct_Apply_Link
        );

        if (currentRunHashes.has(urlHash)) {
            continue;
        }

        currentRunHashes.add(urlHash);

        // Already sent in previous runs?
        if (seenHashes.has(urlHash)) {
            continue;
        }

        // ----------------------------------------------------
        // Stack Match
        // ----------------------------------------------------

        const matchedSkills = [
            ...analysis.matchedStrongSkills,
            ...analysis.matchedSupportingSkills,
        ];

        job.Stack_Match =
            [...new Set(matchedSkills)].join(', ');

        job.Match_Score =
            analysis.score;

        job.Priority =
            analysis.priority;

        matchedJobs.push(job);

        seenHashes.add(urlHash);
    }

    // ========================================================
    // SORT BY MATCH SCORE
    // ========================================================

    matchedJobs.sort(
        (a, b) => b.Match_Score - a.Match_Score
    );

    console.log('');
    console.log(
        `🎯 New matching jobs: ${matchedJobs.length}`
    );

    // ========================================================
    // NO RESULTS
    // ========================================================

    if (matchedJobs.length === 0) {
        console.log('');
        console.log(
            'ℹ️ No new matching jobs found.'
        );

        // Still save the state.
        await kvClient.setRecord({
            key: 'SEEN_JOB_HASHES',
            value: Array.from(seenHashes),
        });

        return;
    }

    // ========================================================
    // EXCEL
    // ========================================================

    console.log('');
    console.log(
        '📊 Creating Excel report...'
    );

    const workbook =
        new ExcelJS.Workbook();

    workbook.creator =
        'L2 Job Scraper';

    workbook.created =
        new Date();

    // --------------------------------------------------------
    // Main Sheet
    // --------------------------------------------------------

    const ws =
        workbook.addWorksheet(
            'Matching Jobs',
            {
                views: [
                    {
                        state: 'frozen',
                        ySplit: 1,
                    },
                ],
            }
        );

    ws.columns = [
        {
            header: 'Job Title',
            key: 'Job_Title',
            width: 45,
        },
        {
            header: 'Company',
            key: 'Company',
            width: 28,
        },
        {
            header: 'Source',
            key: 'Source',
            width: 15,
        },
        {
            header: 'Date Posted',
            key: 'Date_Posted',
            width: 16,
        },
        {
            header: 'Location',
            key: 'Location',
            width: 25,
        },
        {
            header: 'Experience',
            key: 'Experience',
            width: 18,
        },
        {
            header: 'Stack Match',
            key: 'Stack_Match',
            width: 60,
        },
        {
            header: 'Match Score',
            key: 'Match_Score',
            width: 15,
        },
        {
            header: 'Priority',
            key: 'Priority',
            width: 15,
        },
        {
            header: 'Direct Apply Link',
            key: 'Direct_Apply_Link',
            width: 60,
        },
    ];

    // --------------------------------------------------------
    // Header styling
    // --------------------------------------------------------

    const headerRow =
        ws.getRow(1);

    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
            argb: 'FF1F4E78',
        },
    };

    headerRow.font = {
        color: {
            argb: 'FFFFFFFF',
        },
        bold: true,
    };

    headerRow.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
    };

    // --------------------------------------------------------
    // Add jobs
    // --------------------------------------------------------

    matchedJobs.forEach(job => {
        const row =
            ws.addRow(job);

        row.getCell(
            'Date_Posted'
        ).numFmt =
            'dd-mmm-yyyy';

        // Apply link
        const linkCell =
            row.getCell(
                'Direct_Apply_Link'
            );

        if (
            linkCell.value &&
            String(linkCell.value)
                .startsWith('http')
        ) {
            linkCell.value = {
                text: job.Direct_Apply_Link,
                hyperlink:
                    job.Direct_Apply_Link,
            };

            linkCell.font = {
                color: {
                    argb: 'FF0563C1',
                },
                underline: true,
            };
        }

        row.eachCell(cell => {
            cell.alignment = {
                vertical: 'top',
                wrapText: true,
            };
        });
    });

    // --------------------------------------------------------
    // Excel table
    // --------------------------------------------------------

    ws.addTable({
        name: 'L2MatchingJobs',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        style: {
            theme: 'TableStyleMedium2',
            showRowStripes: true,
        },
        columns: ws.columns.map(
            column => ({
                name: column.header,
                filterButton: true,
            })
        ),
        rows: matchedJobs.map(
            job => [
                job.Job_Title,
                job.Company,
                job.Source,
                job.Date_Posted,
                job.Location,
                job.Experience,
                job.Stack_Match,
                job.Match_Score,
                job.Priority,
                job.Direct_Apply_Link,
            ]
        ),
    });

    // ========================================================
    // SUMMARY SHEET
    // ========================================================

    const summary =
        workbook.addWorksheet(
            'Search Summary'
        );

    summary.columns = [
        {
            header: 'Metric',
            key: 'metric',
            width: 35,
        },
        {
            header: 'Value',
            key: 'value',
            width: 50,
        },
    ];

    summary.addRows([
        {
            metric: 'Search Keywords',
            value: SEARCH_KEYWORDS,
        },
        {
            metric: 'Locations',
            value: TARGET_LOCATIONS.join(', '),
        },
        {
            metric: 'Search Window',
            value: `${SEARCH_DAYS} days`,
        },
        {
            metric: 'Experience',
            value: `${EXPERIENCE_MIN}-${EXPERIENCE_MAX} years`,
        },
        {
            metric: 'Raw Jobs Collected',
            value: allJobs.length,
        },
        {
            metric: 'New Matching Jobs',
            value: matchedJobs.length,
        },
        {
            metric: 'Successful Scraper Runs',
            value: successfulRuns,
        },
        {
            metric: 'Generated',
            value: new Date(),
        },
    ]);

    summary.getRow(1).font = {
        bold: true,
    };

    // ========================================================
    // SAVE FILE
    // ========================================================

    const fileName =
        'L2_Deep_Job_Search.xlsx';

    await workbook.xlsx.writeFile(
        fileName
    );

    console.log('');
    console.log(
        `✅ Excel created: ${fileName}`
    );

    // ========================================================
    // SAVE SEEN JOBS
    // ========================================================

    await kvClient.setRecord({
        key: 'SEEN_JOB_HASHES',
        value: Array.from(seenHashes),
    });

    console.log(
        '💾 Job history saved to Apify Key-Value Store.'
    );

    // ========================================================
    // EMAIL
    // ========================================================

    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_TO || process.env.EMAIL_USER,

        subject:
            `🚀 L2 Job Alert: ${matchedJobs.length} New Matches`,

        text:
            `L2 Job Search completed.

Locations:
${TARGET_LOCATIONS.join(', ')}

Portals:
LinkedIn, Naukri, Hirist, Indeed

Search:
${SEARCH_KEYWORDS}

New matching jobs:
${matchedJobs.length}

The detailed Excel report is attached.`,

        attachments: [
            {
                filename: fileName,
                path: `./${fileName}`,
            },
        ],
    };

    try {
        const info =
            await transporter.sendMail(
                mailOptions
            );

        console.log('');
        console.log(
            `📧 Email sent successfully: ${info.messageId}`
        );
    } catch (error) {
        console.error('');
        console.error(
            '❌ Email failed:',
            error.message
        );
    }

    console.log('');
    console.log(
        '======================================================'
    );
    console.log(
        '✅ JOB SEARCH COMPLETED'
    );
    console.log(
        '======================================================'
    );
}

// ============================================================
// START
// ============================================================

aggregateAndFormatJobs()
    .catch(error => {
        console.error('');
        console.error(
            '💥 FATAL ERROR:',
            error
        );

        process.exitCode = 1;
    });