// 1. MUST BE AT THE VERY TOP
require('dotenv').config();

const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');

const dns = require('dns');
if (dns.setDefaultResultOrder) {
  dns.setDefaultResultOrder('ipv4first'); 
}

const cloudinary = require('cloudinary').v2;
const { CloudinaryStorage } = require('multer-storage-cloudinary');

const app = express();
const PORT = process.env.PORT || 5000;

// --- CORS & MIDDLEWARE ---
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.options('*', cors());
app.use(express.json());

// --- ROOT HEALTH CHECK ROUTE ---
app.get('/', (req, res) => {
  res.send('Cohort Backend API is live and running!');
});

// --- MONGODB CONNECTION ---
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/cohort_db';
mongoose.connect(MONGO_URI)
  .then(() => console.log('MongoDB Database Connected Successfully.'))
  .catch(err => console.error('MongoDB Connection Breakdown Error:', err));

// --- PRE-LOADED STUDENT DATABASE ---
const PRELOADED_STUDENTS = [
  {
    rollNumber: "25NG5A0501",
    surname: "Annam",
    lastName: "Sruthi",
    emailId: "annamsruthi07@gmail.com",
    mobileNumber: "9059853496"
  },
  {
    rollNumber: "25NG5A0502",
    surname: "A.V.S.D",
    lastName: "Lokesh",
    emailId: "lokeshllp123@gmail.com",
    mobileNumber: "7416549508"
  },
  {
    rollNumber: "25NG5A0503",
    surname: "Boppana",
    lastName: "Revani",
    emailId: "brevani588@gmail.com",
    mobileNumber: "7702265706"
  },
  {
    rollNumber: "25NG5A0504",
    surname: "Ch",
    lastName: "Kotesh",
    emailId: "chkotesh52@gmail.com",
    mobileNumber: "8019763443"
  },
  {
    rollNumber: "25NG5A0505",
    surname: "Ch",
    lastName: "Dharani",
    emailId: "chdharani58@gmail.com",
    mobileNumber: "9110594011"
  },
  {
    rollNumber: "25NG5A0506",
    surname: "Gadelli",
    lastName: "Chaitanya Prasad",
    emailId: "21351cm012@gmail.com",
    mobileNumber: "9666598898"
  },
  {
    rollNumber: "25NG5A0507",
    surname: "Guduru",
    lastName: "Radha Krishna Phanedhra Kumar",
    emailId: "gudururadhakrishna109@gmail.com",
    mobileNumber: "9494760886"
  },
  {
    rollNumber: "25NG5A0508",
    surname: "Gunja",
    lastName: "Jaya Krishna",
    emailId: "gjayakrishnaur25@gmail.com",
    mobileNumber: "9676270846"
  },
  {
    rollNumber: "25NG5A0509",
    surname: "K",
    lastName: "Charan",
    emailId: "charankambampati1@gmail.com",
    mobileNumber: "7993016355"
  },
  {
    rollNumber: "25NG5A0510",
    surname: "Kandlavath",
    lastName: "Yavana Naik",
    emailId: "ramnaik7452@gmail.com",
    mobileNumber: "8309845863"
  },
  {
    rollNumber: "25NG5A0511",
    surname: "Katuri",
    lastName: "Saranya",
    emailId: "saranyakaturi4@gmail.com",
    mobileNumber: "7093778168"
  },
  {
    rollNumber: "25NG5A0512",
    surname: "Kondeti",
    lastName: "Pravalikka",
    emailId: "pravallikakondeti633@gmail.com",
    mobileNumber: "8639166527"
  },
  {
    rollNumber: "25NG5A0513",
    surname: "Korlapudi",
    lastName: "Naga Satya Giridhar",
    emailId: "dhargiri897@gmail.com",
    mobileNumber: "8309591876"
  },
  {
    rollNumber: "25NG5A0514",
    surname: "Kuthadi",
    lastName: "Maharshi Raghava",
    emailId: "raghavakuthadi@gmail.com",
    mobileNumber: "9052466238"
  },
  {
    rollNumber: "25NG5A0515",
    surname: "Manda",
    lastName: "Harshavardhan",
    emailId: "harshavardhanmanda8@gmail.com",
    mobileNumber: "7569750503"
  },
  {
    rollNumber: "25NG5A0516",
    surname: "Manem",
    lastName: "Hemanth",
    emailId: "hemanthmanem62@gmail.com",
    mobileNumber: "7386670155"
  },
  {
    rollNumber: "25NG5A0517",
    surname: "M",
    lastName: "Harika",
    emailId: "harikametla27@gmail.com",
    mobileNumber: "8919936993"
  },
  {
    rollNumber: "25NG5A0518",
    surname: "M",
    lastName: "Lavanya",
    emailId: "mukkulavanya76@gmail.com",
    mobileNumber: "7093579654"
  },
  {
    rollNumber: "25NG5A0519",
    surname: "Muppa",
    lastName: "Aksay Kumar",
    emailId: "jyothimuppa6@gmail.com",
    mobileNumber: "7270044999"
  },
  {
    rollNumber: "25NG5A0520",
    surname: "Natta",
    lastName: "Harishitha",
    emailId: "harshithanatta@gmail.com",
    mobileNumber: "7981170301"
  },
  {
    rollNumber: "25NG5A0521",
    surname: "Rangisetti",
    lastName: "Laskhmi Lavanya",
    emailId: "lavanyarangisetti689@gmail.com",
    mobileNumber: "7569153754"
  },
  {
    rollNumber: "25NG5A0522",
    surname: "Singareddy",
    lastName: "Yaswanth",
    emailId: "yaswanthsingareddy86@gmail.com",
    mobileNumber: "7032128488"
  },
  {
    rollNumber: "25NG5A0523",
    surname: "Somarouthu",
    lastName: "Siva Naga Raju",
    emailId: "somaroutusivanagaraju@gmail.com",
    mobileNumber: "9346131588"
  },
  {
    rollNumber: "25NG5A0524",
    surname: "Yashwanth",
    lastName: "Sunkara",
    emailId: "sunkarayeswanth42@gmail.com",
    mobileNumber: "7013184893"
  },
  {
    rollNumber: "25NG5A0525",
    surname: "Tripuraneni",
    lastName: "Prasannasai",
    emailId: "learnify.prasannasai@gmail.com",
    mobileNumber: "8309419511"
  },
  {
    rollNumber: "25NG5A0526",
    surname: "Vaka",
    lastName: "Komala Devi",
    emailId: "komaladevi2005@gmail.com",
    mobileNumber: "7386772939"
  },
  {
    rollNumber: "25NG5A0527",
    surname: "Veerla",
    lastName: "Manoj",
    emailId: "manojveerla01@gmail.com",
    mobileNumber: "8897294974"
  },
  {
    rollNumber: "25NG5A0528",
    surname: "Yandrapu",
    lastName: "Rohith Kumar",
    emailId: "rk2007230@gmail.com",
    mobileNumber: "8309673938"
  },
  {
    rollNumber: "25NG5A0529",
    surname: "Veerla",
    lastName: "Sai Keerthana",
    emailId: "veerlasaikeerthana130@gmail.com",
    mobileNumber: "8466802922"
  }
];

// --- MONGODB SCHEMA ---
const cohortCertificateSchema = new mongoose.Schema({
  surname: { type: String, required: true, trim: true },
  lastName: { type: String, trim: true, default: null },
  rollNumber: { type: String, required: true, uppercase: true, trim: true },
  emailId: { type: String, required: true, lowercase: true, trim: true },
  mobileNumber: { type: String, required: true, trim: true },
  courseName: { type: String, required: true, trim: true },
  fileUrl: { type: String, required: true }
}, { timestamps: true });

const CohortCertificate = mongoose.model('CohortCertificate', cohortCertificateSchema);

// --- CLOUDINARY CONFIGURATION ---
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME || '',
  api_key: process.env.CLOUDINARY_API_KEY || '',
  api_secret: process.env.CLOUDINARY_API_SECRET || ''
});

const storage = new CloudinaryStorage({
  cloudinary: cloudinary,
  params: async (req, file) => {
    return {
      folder: 'cohort_17_certificates',
      allowed_formats: ['jpg', 'png', 'jpeg', 'pdf'],
      resource_type: 'auto'
    };
  },
});
const upload = multer({ storage: storage });

const ADMIN_SECRET_KEY = process.env.ADMIN_SECRET_KEY || "admin@urcet"; 

// --- ROUTES ---

// 1. Fetch Roll Numbers
app.get('/api/roll-numbers', async (req, res) => {
  try {
    const rollList = PRELOADED_STUDENTS.map(s => s.rollNumber);
    return res.json({ success: true, rollNumbers: rollList });
  } catch (error) {
    console.error("Roll Numbers Error:", error);
    return res.status(500).json({ success: false, message: "Error fetching roll numbers." });
  }
});

// 2. Student Lookup by Roll Number
app.get('/api/student-lookup/:rollNumber', async (req, res) => {
  try {
    const targetRoll = req.params.rollNumber.toUpperCase();
    const student = PRELOADED_STUDENTS.find(s => s.rollNumber.toUpperCase() === targetRoll);

    if (!student) {
      return res.status(404).json({ success: false, message: "Roll Number not found in student directory." });
    }

    return res.json({ success: true, student });
  } catch (error) {
    console.error("Lookup Error:", error);
    return res.status(500).json({ success: false, message: "Error performing student lookup." });
  }
});

// Helper Handler for Form Submissions
const handleClearanceSubmission = async (req, res) => {
  try {
    const { surname, lastName, rollNumber, emailId, mobileNumber, courseName } = req.body;
    
    if (!surname || !rollNumber || !emailId || !mobileNumber || !courseName || !req.file) {
      return res.status(400).json({ success: false, message: "Please fill all required inputs and upload files." });
    }

    const fileDestination = req.file.path; // Hosted Cloudinary URL

    const record = await CohortCertificate.create({
      surname: surname.trim(),
      lastName: lastName ? lastName.trim() : null,
      rollNumber: rollNumber.toUpperCase().trim(),
      emailId: emailId.trim().toLowerCase(),
      mobileNumber: mobileNumber.trim(),
      courseName: courseName.trim(),
      fileUrl: fileDestination
    });

    console.log(`[Clearance Submitted] Roll: ${rollNumber}, File: ${fileDestination}`);

    return res.status(201).json({ 
      success: true, 
      message: "Records updated successfully!", 
      fileUrl: fileDestination, 
      record 
    });
  } catch (error) {
    console.error("Submit Error:", error);
    return res.status(500).json({ success: false, message: "Database server failure processing upload." });
  }
};

// 3. Primary Endpoint called by forms.html
app.post('/api/submit-clearance', upload.single('attachedCertificate'), handleClearanceSubmission);

// Alias Endpoint for legacy calls
app.post('/api/submit-cohort', upload.single('attachedCertificate'), handleClearanceSubmission);

// 4. Secure Cohort List Dashboard
app.post('/api/secure-cohort-list', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) {
      return res.status(401).json({ success: false, message: "Unauthorized dashboard entry blocked." });
    }
    const data = await CohortCertificate.find().sort({ createdAt: -1 });
    return res.json({ success: true, data: data });
  } catch (error) { 
    return res.status(500).json({ success: false, error: "Database retrieval crash." }); 
  }
});

// 5. Delete Record
app.delete('/api/cohort-records/:id', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) {
      return res.status(401).json({ success: false, message: "Unauthorized deletion attempt blocked." });
    }
    
    const deletedRecord = await CohortCertificate.findByIdAndDelete(req.params.id);
    if (!deletedRecord) {
      return res.status(404).json({ success: false, message: "Record not found." });
    }
    
    return res.json({ success: true });
  } catch (error) { 
    return res.status(500).json({ success: false, message: "Database deletion crash." }); 
  }
});

// 6. Download Excel
app.post('/api/download-cohort-excel', async (req, res) => {
  try {
    const { adminPassword } = req.body;
    if (adminPassword !== ADMIN_SECRET_KEY) return res.status(401).send("Unauthorized.");

    const records = await CohortCertificate.find().sort({ createdAt: -1 });
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Cohort 17 Certifications');

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
        fileUrl: item.fileUrl,
        createdAt: new Date(item.createdAt).toLocaleString()
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Cohort17_Certificate_Records.xlsx');
    await workbook.xlsx.write(res);
    return res.end();
  } catch (error) { 
    console.error("Excel Error:", error);
    return res.status(500).send("Excel generation failure."); 
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
