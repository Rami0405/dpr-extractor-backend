const express = require('express');
const multer = require('multer');
const app = express();
const upload = multer({ dest: 'uploads/' });
const cors = require('cors');
const path = require('path');
const dprController = require('./controllers/dprController');

app.use(cors()); // Enable CORS for all routes

app.post('/api/upload', upload.array('files'), dprController);

app.listen(3000, () => {
    console.log('Server is running on port 3000');
});
