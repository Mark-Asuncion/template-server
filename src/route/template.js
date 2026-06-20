const useDocxReportGenerator = require("../engine/template-engine.js")
/**
 * @param {import('express').Request} req
 * @param {import('express').Response} req
 */
function template_route(req, res) {
    const data = req.body;
    const path = useDocxReportGenerator(data);

    res.send(path);
}

module.exports = template_route;
