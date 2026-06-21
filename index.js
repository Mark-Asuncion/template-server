const express = require("express");
const { join, dirname } = require("path");
const { config } = require("dotenv");
const { fileURLToPath } = require("url");
const template_route = require("./src/route/template.js")

config();

// const __filename = fileURLToPath(import.meta.url);
// const __dirname = dirname(__filename);

const app = express();
const cors = require("cors");
const PORT = process.env.PORT || 3000;
app.locals.XAPIKEY = process.env.XAPIKEY || "";
if (app.locals.XAPIKEY.length == 0) {
    throw new Error("XAPIKEY env is not set");
}
const origins = (process.env.AllowedOrigins || "").split(",");

console.log("origins", origins);
const public_dir = join(__dirname, "public");
const _cors = cors({
  origin: origins,
  methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
  allowedHeaders: "*"
});
app.use(_cors);
app.use(express.json())
app.use("/public/", express.static(public_dir));

/**
* @param {express.Request} req
* @param {express.Response} res
*/
function is_authorize(req, res, next) {
    let k = req.get("X-API-Key");
    if (k !== app.locals.XAPIKEY) {
        return res.sendStatus(401);
    }
    next();
}

app.get("/", is_authorize, (_, res) => {
  res.send("ok");
});

app.post("/template/", is_authorize, template_route);

// app.post("/template/")

app.listen(PORT, () => {
    console.debug("locals", app.locals)
    console.log(`Server running at http://localhost:${PORT}`);
});
