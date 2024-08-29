const express = require('express');
const cors = require('cors'); // Import CORS
const app = express();
const port = 3000;

app.use(cors()); // Enable CORS
app.use(express.json());

app.post('/login', (req, res) => {
    const { email, senha } = req.body;
    if (email === 'user@example.com' && senha === 'password123') {
        res.json({ message: 'Login successful!', status: 'success' });
    } else {
        res.json({ message: 'Invalid credentials', status: 'fail' });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}/`);
});

