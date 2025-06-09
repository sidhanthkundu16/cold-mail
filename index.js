const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const fs = require("fs");

// Load Excel file
const workbook = XLSX.readFile("companies.xlsx");
const sheetName = workbook.SheetNames[0];
const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
console.log(`Loaded ${data.length} entries from Excel.`);

data.forEach((row, i) => {
    console.log(`Row ${i + 1}: Company - ${row.company_name}, Email - ${row.email}`);
});
// Email configuration
const senderEmail = "sidhanthkundutech@gmail.com";
const senderPassword = "";

const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: senderEmail,
        pass: senderPassword,
    },
});

// Email content template
function createEmailBody(company) {
    return `
Hi ${company} Team,

I'm Sidhanth Kundu, a Software Developer with 2.5 years of experience building scalable systems using React.js, Angular, Node.js (NestJS), MongoDB, and PostgreSQL. I've worked across both fast-moving unicorn startups  and global enterprises—delivering fullstack systems that are as robust as they are efficient.

What excites me about ${company} is your bold approach to innovation. I’m eager to contribute to that momentum by designing systems that perform under pressure and scale with confidence.

I've attached my resume and would welcome the opportunity to connect if there’s a potential fit.

Best regards,  
Sidhanth Kundu  
+919812438502 
https://www.linkedin.com/in/sidhanthkundu/
`;
}

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Main sender loop with delay
async function sendEmails() {
    let limit = 0;
    for (const row of data) {
        const mailOptions = {
            from: senderEmail,
            to: row.email,
            subject: `Software Developer | Immediate Joiner | 2.5 Years Exp. | Excited to Build at ${row.company_name}`,
            text: createEmailBody(row.company_name),
            attachments: [
                {
                    filename: "Resume.pdf",
                    path: "./Resume.pdf",
                },
            ],
        };

        try {
            await transporter.sendMail(mailOptions);
            console.log(`✅ Email sent to ${row.email} (${row.company_name})`);
        } catch (err) {
            console.error(`❌ Failed to send to ${row.email}:`, err.message);
        }
        limit = limit + 1;

        if (limit === 100) {
            break;
        }

        await delay(5000);
    }
}

sendEmails();