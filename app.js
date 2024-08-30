const express = require('express');
const cors = require('cors');
const app = express();
const port = process.env.PORT || 3000;

app.use(cors({ origin: '*' }));
app.use(express.json());

app.get('/login', (req, res) => {
    res.send('Server is up and running.');
});

// app.post('/login', (req, res) => {
//     const { email, senha } = req.body;
//     if (email === 'lucasmac31@hotmail.com' && senha === '123') {
//         res.json({ message: 'Login successful!', status: 'success' });
//     } else {
//         res.json({ message: 'Invalid credentials', status: 'fail' });
//     }
// });

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}/`);
});

