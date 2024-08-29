const express = require('express');
const cors = require('cors');
const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());

app.post('/login', (req, res) => {
    const { email, senha } = req.body;
    if (email === 'lucasmac31@hotmail.com' && senha === '123') {
        res.json({ message: 'Login successful!', status: 'success' });
    } else {
        res.json({ message: 'Invalid credentials', status: 'fail' });
    }
});

app.listen(port, () => {
    console.log("Server running.");
});

