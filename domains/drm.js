'use strict';

function createDrmHandler() {
    return function getPlayReadyAuthorizationUrl(req, res) {
        const token = req.query.token || '';
        const customData = req.query.CustomData || '';
        const response = 'p1=5&p2=&p3=&p4=1&p5=0&p6=1&p7=0&p8=0' + '&token=' + encodeURIComponent(token) + '&CustomData=' + encodeURIComponent(customData);
        res.set('Content-Type', 'text/html');
        res.status(200).send(response);
    };
}

module.exports = { createDrmHandler };
