import 'dotenv/config';
import { ApifyClient } from 'apify-client';
import ExcelJS from 'exceljs';
import nodemailer from 'nodemailer';
import crypto from 'crypto';

const client = new ApifyClient({
    token: process.env.APIFY_API_TOKEN, 
});

const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER, 
        pass: process.env.EMAIL_PASS
    }
});

const generateMD5 = (string) => crypto.createHash('md5').update(string).digest('hex');

async function aggregateAndFormatJobs() {
    console.log('🚀 Launching multi-location parallel scrapers...');
    const allJobs = [];
    
    const coreKeywords = "Full Stack Engineer React Node AWS";
    // 🎯 Expanded Location Array
    const targetLocations = ["Chennai", "Bengaluru", "Hyderabad", "Trichy", "Coimbatore"];
    const past21Days = 21; 

    const store = await client.keyValueStores().getOrCreate('l2-job-scraper-state');
    const kvClient = client.keyValueStore(store.id);
    const seenRecord = await kvClient.getRecord('SEEN_JOB_HASHES');
    const seenHashes = new Set(seenRecord ? seenRecord.value : []);

    // 1. Dynamically Build Scraper Promises
    const scrapeTasks = [];
    
    targetLocations.forEach(loc => {
        // Google Jobs
        scrapeTasks.push(client.actor("orgupdate/google-jobs-scraper").call({ includeKeyword: coreKeywords, locationName: `${loc}, India`, countryName: "india", datePosted: "month" })
            .then(r => client.dataset(r.defaultDatasetId).listItems().then(d => ({ portal: 'Google', loc, items: d.items }))));
        
        // LinkedIn
        scrapeTasks.push(client.actor("valig/linkedin-jobs-scraper").call({ keywords: coreKeywords, location: loc, datePosted: "r1814400", limit: 30 })
            .then(r => client.dataset(r.defaultDatasetId).listItems().then(d => ({ portal: 'LinkedIn', loc, items: d.items }))));
        
        // Naukri
        scrapeTasks.push(client.actor("epicscrapers/naukri-scraper").call({ keyword: coreKeywords, location: loc, experienceMin: 2, experienceMax: 4, postedWithinDays: past21Days })
            .then(r => client.dataset(r.defaultDatasetId).listItems().then(d => ({ portal: 'Naukri', loc, items: d.items }))));
        
        // Indeed
        scrapeTasks.push(client.actor("automation-lab/indeed-scraper").call({ query: coreKeywords, location: loc, country: "IN", maxDaysOld: past21Days, maxItems: 30 })
            .then(r => client.dataset(r.defaultDatasetId).listItems().then(d => ({ portal: 'Indeed', loc, items: d.items }))));
        
        // Hirist
        scrapeTasks.push(client.actor("hipersoft/hirist-scraper").call({ queries: [coreKeywords], location: loc, maxDaysOld: past21Days })
            .then(r => client.dataset(r.defaultDatasetId).listItems().then(d => ({ portal: 'Hirist', loc, items: d.items }))));
    });

    console.log(`📡 Queued \({scrapeTasks.length} scraper tasks across\){targetLocations.length} locations...`);
    const results = await Promise.allSettled(scrapeTasks);

    // 2. Map Results Dynamically
    const mapJob = (job, portal, loc, titleKey, companyKey, dateKey, linkKey, descKey) => {
        if (!job || !job[linkKey]) return;
        allJobs.push({
            Job_Title: job[titleKey] || 'N/A',
            Company: job[companyKey] || 'N/A',
            Source: portal,
            Date_Posted: job[dateKey] ? new Date(job[dateKey]) : new Date(),
            Location: job.location || loc, // Favor specific scraped location, fallback to search city
            Experience: "2-4 years (Estimated)", 
            Stack_Match: "React, Node.js, AWS", 
            Priority: "High", 
            Direct_Apply_Link: job[linkKey],
            Career_Page: "Search company careers", 
            Description_Text: job[descKey] || '' 
        });
    };

    results.forEach(res => {
        if (res.status === 'fulfilled') {
            const { portal, loc, items } = res.value;
            items.forEach(j => {
                if (portal === 'Google') mapJob(j, 'Google', loc, 'title', 'companyName', 'datePosted', 'applyLink', 'description');
                else if (portal === 'LinkedIn') mapJob(j, 'LinkedIn', loc, 'title', 'companyName', 'postedAt', 'url', 'description');
                else if (portal === 'Naukri') mapJob(j, 'Naukri', loc, 'title', 'companyName', 'postedDate', 'jobUrl', 'jobDescription');
                else if (portal === 'Indeed') mapJob(j, 'Indeed', loc, 'jobTitle', 'company', 'date', 'url', 'description');
                else if (portal === 'Hirist') mapJob(j, 'Hirist', loc, 'title', 'company', 'date', 'url', 'description');
            });
        }
    });

    // 3. Relaxed Regex & Deduplication
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - past21Days);
    
    // 🎯 Expanded regex to catch React OR Node OR Next roles to increase volume
    const regexFilter = /react|node\.js|nodejs|next\.js|nextjs/i;
    const primaryMatches = [];

    for (const job of allJobs) {
        const isRecent = job.Date_Posted >= cutoffDate;
        const matchesStack = regexFilter.test(job.Job_Title) || regexFilter.test(job.Description_Text);
        
        if (isRecent && matchesStack) {
            const urlHash = generateMD5(job.Direct_Apply_Link);
            if (!seenHashes.has(urlHash)) {
                seenHashes.add(urlHash);
                primaryMatches.push(job);
            }
        }
    }

    if (primaryMatches.length === 0) {
        console.log('No new jobs matched the expanded criteria. Exiting.');
        return;
    }

    // 4. Build Styled Excel Workbook
    console.log(`🎨 Formatting ${primaryMatches.length} jobs into Excel...`);
    const workbook = new ExcelJS.Workbook();
    
    const ws = workbook.addWorksheet('Strict Matches', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [
        { header: 'Job Title', key: 'Job_Title', width: 48 },
        { header: 'Company', key: 'Company', width: 28 },
        { header: 'Source', key: 'Source', width: 18 },
        { header: 'Date Posted', key: 'Date_Posted', width: 14 },
        { header: 'Location', key: 'Location', width: 26 },
        { header: 'Experience', key: 'Experience', width: 18 },
        { header: 'Stack Match', key: 'Stack_Match', width: 65 },
        { header: 'Priority', key: 'Priority', width: 14 },
        { header: 'Direct Apply Link', key: 'Direct_Apply_Link', width: 58 },
        { header: 'Career Page Apply Link', key: 'Career_Page', width: 50 }
    ];

    const headerRow = ws.getRow(1);
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    headerRow.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    primaryMatches.forEach(job => {
        const row = ws.addRow(job);
        row.getCell('Date_Posted').numFmt = 'dd-mmm-yyyy';
        
        const linkCell = row.getCell('Direct_Apply_Link');
        if (linkCell.value && linkCell.value.toString().startsWith('http')) {
            linkCell.value = { text: job.Direct_Apply_Link, hyperlink: job.Direct_Apply_Link };
            linkCell.font = { color: { argb: 'FF0563C1' }, underline: true };
        }
        
        row.eachCell(cell => cell.alignment = { vertical: 'top', wrapText: true });
    });

    ws.addTable({
        name: 'StrictJobMatches',
        ref: 'A1',
        headerRow: true,
        totalsRow: false,
        style: { theme: 'TableStyleMedium2', showRowStripes: true },
        columns: ws.columns.map(c => ({ name: c.header, filterButton: true })),
        rows: primaryMatches.map(j => [j.Job_Title, j.Company, j.Source, j.Date_Posted, j.Location, j.Experience, j.Stack_Match, j.Priority, j.Direct_Apply_Link, j.Career_Page])
    });

    // 5. Save and Email
    const fileName = 'Formatted_L2_Jobs_Multi_City.xlsx';
    await workbook.xlsx.writeFile(fileName);
    await kvClient.setRecord({ key: 'SEEN_JOB_HASHES', value: Array.from(seenHashes) });

    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: 'aruljyothi0202@gmail.com',
        subject: `🚀 L2 Full-Stack Alert: ${primaryMatches.length} Matches Found`,
        text: `Searched across Chennai, Bengaluru, Hyderabad, Trichy, and Coimbatore. Found ${primaryMatches.length} net-new matches. The attached Excel is fully styled.`,
        attachments: [{ filename: fileName, path: `./${fileName}` }]
    };

    try {
        const info = await transporter.sendMail(mailOptions);
        console.log(`✅ Email sent: ${info.messageId}`);
    } catch (error) {
        console.error(`❌ Error sending email: ${error}`);
    }
}

aggregateAndFormatJobs().catch(console.error);