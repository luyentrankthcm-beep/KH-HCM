const express = require("express");
const http = require("http");

// Stub auth middleware to bypass login/admin check for this automated run.
const auth = require("./middleware/auth");
auth.requireLogin = (req, res, next) => next();
auth.requireAdmin = (req, res, next) => next();

const chiPhiRouter = require("./routes/chi-phi");

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use("/", chiPhiRouter);

const server = app.listen(0, () => {
  const port = server.address().port;
  const req = http.request(
    {
      host: "127.0.0.1",
      port,
      path: "/chi-phi/quet-ngan-hang",
      method: "POST",
      headers: { "Content-Length": 0 },
    },
    (res) => {
      const location = res.headers.location || "";
      console.log("STATUS:", res.statusCode);
      console.log("LOCATION:", location);
      let qs = "";
      const idx = location.indexOf("?");
      if (idx !== -1) qs = location.slice(idx + 1);
      const params = new URLSearchParams(qs);
      if (params.get("success")) {
        console.log("SUCCESS_MSG:", decodeURIComponent(params.get("success")));
      }
      if (params.get("error")) {
        console.log("ERROR_MSG:", decodeURIComponent(params.get("error")));
      }
      res.resume();
      res.on("end", () => server.close());
    }
  );
  req.end();
});
