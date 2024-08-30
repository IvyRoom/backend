const express = require('express');
const app = express();
const port = process.env.PORT || 3000;
app.listen(port);
app.use(express.json());

app.get('/login', (req, res) => {
    res.send('Surf is up.');
});

