const express = require('express');
const mongoose = require('mongoose'); // Changed from Sequelize to Mongoose
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware Pipeline
app.use(cors());
app.use(express.json());
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// 1. ESTABLISH MONGODB DATABASE CONNECTION LIFECYCLE
// Your .env should have: MONGO_URI=mongodb://localhost:27017/cohort_db or MongoDB Atlas URI
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/cohort_db';

mongoose.connect(MONGO_URI)
  .then(() => console.log('MongoDB Database Connected Successfully.'))
  .catch(err => console.error('MongoDB Connection Breakdown Error:', err));

// 2. DEFINE SYSTEM SCHEMA & MODEL (MongoDB / Mongoose Setup)
const cohortCertificateSchema = new mongoose.Schema({
  surname: { type: String, required: true, trim: true },
  lastName: { type: String, trim: true, default: null },
  rollNumber: { type: String, required: true, uppercase: true, trim: true },
  emailId: { type: String, required: true, lowercase: true, trim: true },
  mobileNumber: { type: String, required: true, trim: true },
  courseName: { type: String, required: true, trim: true },
  fileUrl: { type: String, required: true }
}, { 
  timestamps: true // Automatically creates 'createdAt' and 'updatedAt' fields
});

const CohortCertificate = mongoose.model('CohortCertificate', cohortCertificateSchema);

// 3. NODEMAILER EMAIL TRANSPORT ENGINE CONFIGURATION
const mailTransporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: parseInt(process.env.SMTP_PORT || '465'),
  secure: true,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS
  }
});

// Helper Function: Dispatch HTML email receipts cleanly
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
  } catch (error) {
    console.error("Email Delivery Pipeline Failure:", error);
  }
}

// 4. MULTER REPLICATED FILE STORAGE CONFIGURATION ENGINE
const storage = multer.diskStorage({
  destination: (req, file, cb) => { cb(null, 'uploads/'); },
  filename: (req, file, cb) => { cb(null, `${Date.now()}-${file.originalname}`); }
});
const upload = multer({ storage: storage });

// MASTER SECRET GATEWAY PASSWORD KEY
const ADMIN_SECRET_KEY = "admin@urcet"; 

// 5. API ROUTING CONTROLLERS

// =========================================================================
// POST: Post Cohort Submission Form (MongoDB Integrated)
// =========================================================================
app.post('/api/submit-cohort', upload.single('attachedCertificate'), async (req, res) => {
  try {
    const { surname, lastName, rollNumber, emailId, mobileNumber, courseName } = req.body;
    
    if (!surname || !rollNumber || !emailId || !mobileNumber || !courseName || !req.file) {
      return res.status(400).json({ success: false, message: "Please fill all required inputs and upload files." });
    }

    const fileDestination = `/uploads/${req.file.filename}`;

    // MongoDB Data Insertion using Mongoose
    await CohortCertificate.create({
      surname,
      lastName,
      rollNumber,
      emailId,
      mobileNumber,
      courseName,
      fileUrl: fileDestination
    });

    // BEAUTIFUL PROFESSIONAL RESPONSIVE EMAIL BLOCKS TEMPLATE
    const luxuryMailTemplate = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Cohort 16 Clearances</title>
      </head>
      <body style="margin: 0; padding: 0; background-color: #f4f6f8; font-family: 'Segoe UI', Helvetica, Arial, sans-serif;">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" style="table-layout: fixed; background-color: #f4f6f8;">
          <tr>
            <td align="center" style="padding: 40px 10px;">
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.05);">
                <tr>
                  <td align="center" style="background: linear-gradient(135deg, #6b1023, #8a1f3d); padding: 40px 30px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                      <tr>
                        <td align="center" style="color: #ffffff; font-size: 26px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase;">
                          urcenoteX Registry
                        </td>
                      </tr>
                      <tr>
                        <td align="center" style="color: #ffcad4; font-size: 14px; font-weight: 600; padding-top: 5px; text-transform: uppercase; letter-spacing: 2px;">
                          Cohort 16 Certification Panel
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td style="padding: 40px 35px; background-color: #ffffff;">
                    <p style="margin: 0 0 15px 0; font-size: 18px; font-weight: 700; color: #1a202c;">Dear ${surname} ${lastName || ''},</p>
                    <p style="margin: 0 0 25px 0; font-size: 15px; color: #4a5568; line-height: 1.6;"> Your graduation records clearance request summary loop completed flawlessly. Your final course completion certificate file has been logged securely into our MongoDB Cloud Systems.</p>
                    
                    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f7fafc; border: 1px solid #edf2f7; border-radius: 12px; padding: 20px; margin-bottom: 25px;">
                      <tr>
                        <td colspan="2" style="font-size: 13px; font-weight: 800; color: #6b1023; text-transform: uppercase; letter-spacing: 0.5px; padding-bottom: 12px; border-bottom: 1px solid #e2e8f0;">
                          Profile Registry Attributes
                        </td>
                      </tr>
                      <tr>
                        <td style="padding: 12px 0 6px 0; font-size: 14px; color: #718096; width: 40%;"><strong>Roll Number:</strong></td>
                        <td style="padding: 12px 0 6px 0; font-size: 14px; color: #1a202c; font-weight: 600;">${rollNumber.toUpperCase()}</td>
                      </tr>
                      <tr>
                        <td style="padding: 6px 0; font-size: 14px; color: #718096;"><strong>Registered Course:</strong></td>
                        <td style="padding: 6px 0; font-size: 14px; color: #1a202c; font-weight: 600;">${courseName}</td>
                      </tr>
                      <tr>
                        <td style="padding: 6px 0; font-size: 14px; color: #718096;"><strong>Mobile Reference:</strong></td>
                        <td style="padding: 6px 0; font-size: 14px; color: #1a202c; font-weight: 600;">${mobileNumber}</td>
                      </tr>
                      <tr>
                        <td style="padding: 6px 0; 12px 0; font-size: 14px; color: #718096; border-bottom: 1px solid #edf2f7;"><strong>Timestamp:</strong></td>
                        <td style="padding: 6px 0; 12px 0; font-size: 14px; color: #2b6cb0; font-weight: 600; border-bottom: 1px solid #edf2f7;">${new Date().toLocaleString()}</td>
                      </tr>
                    </table>
                    <p style="margin: 0; font-size: 13px; color: #a0aec0; line-height: 1.5; text-align: center;">This message acts as an automated submission receipt file transaction proof tracking block. Do not attempt direct email replies.</p>
                  </td>
                </tr>
                <tr>
                  <td align="center" style="background-color: #edf2f7; padding: 20px; font-size: 12px; color: #718096; border-top: 1px solid #e2e8f0;">
                    &copy; ${new Date().getFullYear()} urcenoteX System Clusters. All rights reserved.
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

// =========================================================================
// FETCH: Fetch Secure Cohort Submissions List (Sorted by latest)
// =========================================================================
app.post('/api/secure-cohort-list', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) {
      return res.status(401).json({ success: false, message: "Unauthorized dashboard entry blocked." });
    }
    // Mongoose query syntax changed from findAll() to find().sort()
    const data = await CohortCertificate.find().sort({ createdAt: -1 });
    res.json({ success: true, data: data });
  } catch (error) {
    res.status(500).json({ success: false, error: "Database retrieval crash." });
  }
});

// =========================================================================
// DELETE: Delete Cohort data row by Document ID
// =========================================================================
app.delete('/api/cohort-records/:id', async (req, res) => {
  try {
    // Mongoose syntax to delete by MongoDB _id object
    await CohortCertificate.findByIdAndDelete(req.params.id);
    res.json({ success: true });
  } catch (error) { 
    res.status(500).json({ success: false }); 
  }
});

// =========================================================================
// EXCEL DOWNLOAD: Generate & Download Styled Cohort Spreadsheet
// =========================================================================
app.post('/api/download-cohort-excel', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) return res.status(401).send("Unauthorized.");

    const records = await CohortCertificate.find().sort({ createdAt: -1 });
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cohort 16 Certifications');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 28 }, // Increased width for MongoDB Hex ObjectID strings
      { header: 'Surname', key: 'surname', width: 18 },
      { header: 'Last Name', key: 'lastName', width: 18 },
      { header: 'Roll Number', key: 'rollNumber', width: 18 },
      { header: 'Email ID', key: 'emailId', width: 25 },
      { header: 'Mobile Number', key: 'mobileNumber', width: 18 },
      { header: 'Registered Course Name', key: 'courseName', width: 30 },
      { header: 'Certificate Attachment URL', key: 'fileUrl', width: 50 },
      { header: 'Submission Date', key: 'createdAt', width: 22 }
    ];

    worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' }, name: 'Segoe UI' };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '6B1023' } };

    records.forEach(item => {
      worksheet.addRow({
        id: item._id.toString(), // Convert MongoDB ObjectId to string format cleanly
        surname: item.surname,
        lastName: item.lastName || '',
        rollNumber: item.rollNumber,
        emailId: item.emailId,
        mobileNumber: item.mobileNumber,
        courseName: item.courseName,
        fileUrl: `http://localhost:${PORT}${item.fileUrl}`,
        createdAt: new Date(item.createdAt).toLocaleString()
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Cohort16_Certificate_Records.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) { 
    console.error(error);
    res.status(500).send("Excel generation failure."); 
  }
});

app.listen(PORT, () => console.log(`Cohort 16 MongoDB Engine listening on port ${PORT}`));