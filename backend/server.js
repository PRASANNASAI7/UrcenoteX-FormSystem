const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
require('dotenv').config();

// --- CLOUDINARY INTEGRATION LIBRARIES ---
const cloudinary = require('cloudinary').v2;
const { CloudinaryStorage } = require('multer-storage-cloudinary');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware Pipeline
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.options('*', cors());
app.use(express.json());

// 1. ESTABLISH MONGODB DATABASE CONNECTION
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/cohort_db';
mongoose.connect(MONGO_URI)
  .then(() => console.log('MongoDB Database Connected Successfully.'))
  .catch(err => console.error('MongoDB Connection Breakdown Error:', err));

// 2. DEFINE SYSTEM SCHEMA & MODEL
const cohortCertificateSchema = new mongoose.Schema({
  surname: { type: String, required: true, trim: true },
  lastName: { type: String, trim: true, default: null },
  rollNumber: { type: String, required: true, uppercase: true, trim: true },
  emailId: { type: String, required: true, lowercase: true, trim: true },
  mobileNumber: { type: String, required: true, trim: true },
  courseName: { type: String, required: true, trim: true },
  fileUrl: { type: String, required: true } // This will now hold Cloudinary Live URL (https://res.cloudinary.com/...)
}, { timestamps: true });

const CohortCertificate = mongoose.model('CohortCertificate', cohortCertificateSchema);

// 3. CLOUDINARY CONFIGURATION ENGINE
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

// Configure Multer to upload straight to Cloudinary Vault
const storage = new CloudinaryStorage({
  cloudinary: cloudinary,
  params: {
    folder: 'cohort_16_certificates', // Cloudinary లో క్రియేట్ అయ్యే ఫోల్డర్ పేరు
    allowed_formats: ['jpg', 'png', 'jpeg', 'pdf'], // అనుమతించబడే ఫైల్ ఫార్మాట్లు
    resource_type: 'auto' // Handle both images and PDFs automatically
  },
});
const upload = multer({ storage: storage });

// NODEMAILER TRANSPORT
const mailTransporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: parseInt(process.env.SMTP_PORT || '465'),
  secure: true,
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
});

async function dispatchConfirmationEmail(targetEmail, subjectText, htmlBodyContent) {
  try {
    const mailOptions = {
      from: `"Cohort 16 Administration" <${process.env.SMTP_USER}>`,
      to: targetEmail,
      subject: subjectText,
      html: htmlBodyContent
    };
    await mailTransporter.sendMail(mailOptions);
    console.log(`Auto-Notification delivered to: ${targetEmail}`);
  } catch (error) { console.error("Email Delivery Pipeline Failure:", error); }
}

const ADMIN_SECRET_KEY = "admin@urcet"; 

// =========================================================================
// API ROUTE: Post Cohort Submission Form (Cloudinary Integrated)
// =========================================================================
app.post('/api/submit-cohort', upload.single('attachedCertificate'), async (req, res) => {
  try {
    const { surname, lastName, rollNumber, emailId, mobileNumber, courseName } = req.body;
    
    // Cloudinary upload details will be inside req.file.path instead of req.file.filename
    if (!surname || !rollNumber || !emailId || !mobileNumber || !courseName || !req.file) {
      return res.status(400).json({ success: false, message: "Please fill all required inputs and upload files." });
    }

    // req.file.path gives the secure cloud URL (https://res.cloudinary.com/...)
    const fileDestination = req.file.path; 

    await CohortCertificate.create({
      surname: surname.trim(),
      lastName: lastName ? lastName.trim() : null,
      rollNumber: rollNumber.toUpperCase().trim(),
      emailId: emailId.trim().toLowerCase(),
      mobileNumber: mobileNumber.trim(),
      courseName: courseName.trim(),
      fileUrl: fileDestination
    });

    // EMAIL TEMPLATE BLOCK
    const luxuryMailTemplate = `
      <!DOCTYPE html>
      <html>
      <head><meta charset="utf-8"></head>
      <body style="margin:0;padding:0;background-color:#f4f6f8;font-family:'Segoe UI',Arial,sans-serif;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f4f6f8;">
          <tr>
            <td align="center" style="padding:40px 10px;">
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:600px;background-color:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 10px 30px rgba(0,0,0,0.05);">
                <tr>
                  <td align="center" style="background:linear-gradient(135deg, #6b1023, #8a1f3d);padding:40px 30px;color:#ffffff;font-size:24px;font-weight:800;text-transform:uppercase;">
                    urcenoteX Registry
                  </td>
                </tr>
                <tr>
                  <td style="padding:40px 35px;">
                    <p style="font-size:18px;font-weight:700;color:#1a202c;">Dear ${surname} ${lastName || ''},</p>
                    <p style="font-size:15px;color:#4a5568;line-height:1.6;">Your final course completion certificate file has been logged securely into our Cloud Secure Vaults.</p>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f7fafc;border:1px solid #edf2f7;border-radius:12px;padding:20px;margin-bottom:25px;">
                      <tr><td><strong>Roll Number:</strong></td><td>${rollNumber.toUpperCase()}</td></tr>
                      <tr><td><strong>Course:</strong></td><td>${courseName}</td></tr>
                      <tr><td><strong>Certificate Link:</strong></td><td><a href="${fileDestination}" target="_blank">View Uploaded Document</a></td></tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    dispatchConfirmationEmail(emailId.trim().toLowerCase(), "Clearance Receipt: Cohort 16 Certificate Logged", luxuryMailTemplate);

    res.status(201).json({ success: true, message: "Records updated successfully!" });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: "Database server failure loops." });
  }
});

// FETCH SECURE COHORT LIST
app.post('/api/secure-cohort-list', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) {
      return res.status(401).json({ success: false, message: "Unauthorized dashboard entry blocked." });
    }
    const data = await CohortCertificate.find().sort({ createdAt: -1 });
    res.json({ success: true, data: data });
  } catch (error) { res.status(500).json({ success: false, error: "Database retrieval crash." }); }
});

// DELETE ROUTE
app.delete('/api/cohort-records/:id', async (req, res) => {
  try {
    await CohortCertificate.findByIdAndDelete(req.params.id);
    res.json({ success: true });
  } catch (error) { res.status(500).json({ success: false }); }
});

// EXCEL EXPORT (Will now print the live Cloudinary HTTPS URLs)
app.post('/api/download-cohort-excel', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) return res.status(401).send("Unauthorized.");

    const records = await CohortCertificate.find().sort({ createdAt: -1 });
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cohort 16 Certifications');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 28 },
      { header: 'Surname', key: 'surname', width: 18 },
      { header: 'Last Name', key: 'lastName', width: 18 },
      { header: 'Roll Number', key: 'rollNumber', width: 18 },
      { header: 'Email ID', key: 'emailId', width: 25 },
      { header: 'Mobile Number', key: 'mobileNumber', width: 18 },
      { header: 'Registered Course Name', key: 'courseName', width: 30 },
      { header: 'Certificate Attachment URL', key: 'fileUrl', width: 60 },
      { header: 'Submission Date', key: 'createdAt', width: 22 }
    ];

    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Segoe UI' };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6B1023' } };

    records.forEach(item => {
      worksheet.addRow({
        id: item._id.toString(),
        surname: item.surname,
        lastName: item.lastName || '',
        rollNumber: item.rollNumber,
        emailId: item.emailId,
        mobileNumber: item.mobileNumber,
        courseName: item.courseName,
        fileUrl: item.fileUrl, // Direct online link
        createdAt: new Date(item.createdAt).toLocaleString()
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Cohort16_Certificate_Records.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) { res.status(500).send("Excel generation failure."); }
});

app.listen(PORT, () => console.log(`Cohort 16 Cloudinary Engine listening on port ${PORT}`));