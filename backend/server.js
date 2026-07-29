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
  },
   {
    rollNumber: "25NG5A05F1",
    surname: "NERUSU",
    lastName: "BHANUJA",
    emailId: "bhanujanerusu@gmail.com",
    mobileNumber: "6301319934"
  },
  {
    rollNumber: "24NG1A05F2",
    surname: "PAVULURI",
    lastName: "SOWJANYA",
    emailId: "sowjisowjanya2007@gmail.com",
    mobileNumber: "7995258387"
  },
  {
    rollNumber: "24NG1A05F3",
    surname: "PETETI",
    lastName: "LOKSHITHA",
    emailId: "petetilokshitha19@gmail.com",
    mobileNumber: "9704396737"
  },
{
    rollNumber: "24NG1A05F4",
    surname: "PULIGADDA",
    lastName: "DHARANI SRI",
    emailId: "dharanisri94342@gmail.com",
    mobileNumber: "7981507335"
  },
{
    rollNumber: "24NG1A05F5",
    surname: "PUNNA",
    lastName: "JANAKI RAM CHARAN",
    emailId: "punnajanakiramcharan@gmail.com",
    mobileNumber: "9959037531"
  },
{
    rollNumber: "24NG1A05F6",
    surname: "RACHURI",
    lastName: "SRI SAI SIRISHA",
    emailId: "srisaisirisharachuri@gmail.com",
    mobileNumber: "7993258509"
  },
{
    rollNumber: "24NG1A05F7",
    surname: "RAMALA",
    lastName: "SPURTHI",
    emailId: "spoorthiramala@gmail.com",
    mobileNumber: "8074753732"
  },
{
    rollNumber: "24NG1A05F8",
    surname: "SANKARASETTI",
    lastName: "PRATIBHATEJASWI",
    emailId: "pratibhatejaswi752@gmail.com",
    mobileNumber: "8142647999"
  },

{
    rollNumber: "24NG1A05F9",
    surname: "SAYYED",
    lastName: "LATHIBEE",
    emailId: "sdlathibee@gmail.com",
    mobileNumber: "9502562825"
  },
{
    rollNumber: "24NG1A05G0",
    surname: "SHAIK",
    lastName: "HEENA KOUSAR",
    emailId: "afrinseema3@gmail.com",
    mobileNumber: "8499958852"
  },

{
    rollNumber: "24NG1A05G1",
    surname: "SHAIK",
    lastName: "MASUMA",
    emailId: "skmasuma12306@gmail.com",
    mobileNumber: "9390086552"
  },

{
    rollNumber: "24NG1A05G2",
    surname: "SHAIK",
    lastName: "SABEER",
    emailId: "sabeer22006@gmail.com",
    mobileNumber: "9391747269"
  },

{
    rollNumber: "24NG1A05G3",
    surname: "SHAIK",
    lastName: "SHASHMI",
    emailId: "shaikshashmi407@gmail.com",
    mobileNumber: "8985421963"
  },

{
    rollNumber: "24NG1A05G4",
    surname: "SYKAM",
    lastName: "RAKSHAK",
    emailId: "sunnysykam330@gmail.com",
    mobileNumber: "9100087704"
  },

{
    rollNumber: "24NG1A05G5",
    surname: "TANGELLAMUDI",
    lastName: "SWATHI",
    emailId: "swathitangellamudi0@gmail.com",
    mobileNumber: "9573739432"
  },

{
    rollNumber: "24NG1A05G6",
    surname: "TANURI",
    lastName: "SRINIVAS",
    emailId: "srinivastanuri909@gmail.com",
    mobileNumber: "8885410298"
  },

{
    rollNumber: "24NG1A05G7",
    surname: "VEERABATTINA",
    lastName: "GNANA RAJU",
    emailId: "gnanaraju610@gmail.com",
    mobileNumber: "7416835289"
  },

{
    rollNumber: "24NG1A05G8",
    surname: "VELCHURI",
    lastName: "BRAHMA TEJA",
    emailId: "vvelchuribrahmateja@gmail.com",
    mobileNumber: "9391318982"
  },
{
    rollNumber: "24NG1A05G9",
    surname: "VEMULA",
    lastName: "SAI KOWSHTUB",
    emailId: "vemulasaithanu@gmail.com",
    mobileNumber: "6309783808"
  },
 {
    rollNumber: "25NG5A05H0",
    surname: "VEMURI",
    lastName: "LOKESH SAI KUMAR",
    emailId: "lokeshvemuri87@gmail.com",
    mobileNumber: "6301986503"
  },
{
  rollNumber: "24NG1A05H1",
  surname: "ALUBILLI",
  lastName: "SWATHI",
  emailId: "alubilliswathi06@gmail.com",
  mobileNumber: "7569854723"	
},
{
  rollNumber: "24NG1A05H2",
  surname: "ALUGOLU",
  lastName: "DOLA HARINI",
  emailId: "alugoludevi55@gmail.com",
  mobileNumber: "8008469848"
},
{
  rollNumber: "24NG1A05H3",
  surname: "ANNABATHULA",
  lastName: "JAYASRI BHAVANI",
  emailId: "annabathulajayasri@gmail.com",
  mobileNumber: "8374731159"
},
{
  rollNumber: "24NG1A05H5",
  surname: "BAREDDY",
  lastName: "VENKATA SUBBA REDDY",
  emailId: "bareddysubbareddy44@gmail.com",
  mobileNumber: "9573670428"
},
{
  rollNumber: "24NG1A05H6",
  surname: "BEVARA",
  lastName: "VENKATA SIVA BHUVAN DHARMESH",
  emailId: "dharmeshbavara@gmail.com",
  mobileNumber: "9347979542"
},
{
  rollNumber: "24NG1A05H8",
  surname: "BOSETTI",
  lastName: "LAVANYA",
  emailId: "lavanyabosetti@gmail.com",
  mobileNumber: "8688914618"
},
{
  rollNumber: "24NG1A05I1",
  surname: "DUDIGAM",
  lastName: "ANIL",
  emailId: "dudigamanil@gmail.com",
  mobileNumber: "8247213405"
},
{
  rollNumber: "24NG1A05I2",
  surname: "GEDADASU",
  lastName: "ANITHASRI",
  emailId: "anithasrigedadasu@gmail.com",
  mobileNumber: "6300850076"
},
{
  rollNumber: "24NG1A05I3",
  surname: "JOHN",
  lastName: "BLESSY SRI PRIYA",
  emailId: "sripriya66886@gmail.com",
  mobileNumber: "7416938115"
},
{
  rollNumber: "24NG1A05I4",
  surname: "KADIYAM",
  lastName: "GEETHIKA",
  emailId: "geethikakadiyam786@gmail.com",
  mobileNumber: "8897266488"
},
{
  rollNumber: "24NG1A05I5",
  surname: "KARUMANCHI",
  lastName: "ALTHAF",
  emailId: "anwaranwar9164@gmail.com",
  mobileNumber: "6301403744"
},
{
  rollNumber: "24NG1A05I6",
  surname: "KONA",
  lastName: "JANANI",
  emailId: "jananikonas105@gmail.com",
  mobileNumber: "9704089744"
},
{
  rollNumber: "24NG1A05I7",
  surname: "KONDETI",
  lastName: "KARUNA",
  emailId: "karunakondeti7@gmail.com",
  mobileNumber: "9392659417"
},
{
  rollNumber: "24NG1A05I8",
  surname: "MALLIPEDDI",
  lastName: "AKHIL CHOWDARY",
  emailId: "akhil252006@gmail.com",
  mobileNumber: "8367717254"
},
{
  rollNumber: "24NG1A05I9",
  surname: "MANDALI",
  lastName: "JAHNAVI",
  emailId: "mandalijahnavi46@gmail.com",
  mobileNumber: "9502466189"
},
{
  rollNumber: "24NG1A05J0",
  surname: "MARUBOYINA",
  lastName: "HARIKA",
  emailId: "maruboyinaharika2006@gmail.com",
  mobileNumber: "8501963659"
},
{
  rollNumber: "24NG1A05J1",
  surname: "MOHAMMAD",
  lastName: "TURABALI",
  emailId: "mohammadturabali110@gmail.com",
  mobileNumber: "8639236715"
},
{
  rollNumber: "24NG1A05J2",
  surname: "MUNIPALLI",
  lastName: "BHAVYASRI",
  emailId: "munipallibhavyasri@gmail.com",
  mobileNumber: "8555063175"
},
{
  rollNumber: "24NG1A05J3",
  surname: "POLAMARISETTI",
  lastName: "NAGA DURGA POOJITHA",
  emailId: "poojithadurga630@gmail.com",
  mobileNumber: "7675057833"
},
{
  rollNumber: "24NG1A05J5",
  surname: "THOTA",
  lastName: "ASRITHA",
  emailId: "thotaasritha1@gmail.com",
  mobileNumber: "6303778297"
},
{
  rollNumber: "24NG1A05J6",
  surname: "VADDI",
  lastName: "KUSUMA",
  emailId: "vaddisunny0@gmail.com",
  mobileNumber: "9848715059"
},
{
  rollNumber: "24NG1A05J7",
  surname: "VEERLA",
  lastName: "PRASANNA",
  emailId: "prasannaveerla129@gmail.com",
  mobileNumber: "7981978300"
},
{
  rollNumber: "24NG1A05J8",
  surname: "VELIGARAPU",
  lastName: "SAILAJA",
  emailId: "sailajaveligarapu@gmail.com",
  mobileNumber: "6300916619"
},
{
  rollNumber: "24NG1A05J9",
  surname: "YALAMANCHI",
  lastName: "NAGA SAI KRISHNA",
  emailId: "ysaikrishna0222@gmail.com",
  mobileNumber: "9676385751"
},
{
  rollNumber: "24NG1A05K0",
  surname: "YENREDDY",
  lastName: "RAM KISHORE REDDY",
  emailId: "ramkishorereddyyenreddy@gmail.com",
  mobileNumber: "7842073899"
},
{
  rollNumber: "24NG1A05K2",
  surname: "NAKKA",
  lastName: "PUJITHA",
  emailId: "poojitharavikumar33@gmail.com",
  mobileNumber: "8790596043"
},
{
  rollNumber: "24NG1A05K3",
  surname: "ECHARLA",
  lastName: "ADINARAYANA",
  emailId: "echaralaadhi@gmail.com",
  mobileNumber: "6304575899"
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
