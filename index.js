const express = require('express');
const multer = require('multer');
const app = express();
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, '/tmp'); // Use /tmp as the destination directory
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    },
});

const upload = multer({ storage });
const cors = require('cors');
const path = require('path');
const dprController = require('./controllers/dprController');

app.use(cors()); // Enable CORS for all routes

app.post('/api/upload', upload.array('files'), dprController);

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
