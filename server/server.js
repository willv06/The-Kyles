/* ================================
   PhishGuard Backend Server
   ================================ */

   const express = require("express");
   const cors = require("cors");
   const path = require("path");
   
   require("dotenv").config({
       path: path.join(__dirname, "..", ".env")
   });
   
   const app = express();
   
   const PORT = 3001;
   
   
   /* ================================
      Middleware
      ================================ */
   
   app.use(cors());
   app.use(express.json());
   
   
   /* ================================
      Test Route
      ================================ */
   
   app.get("/api/test", (req, res) => {
   
       res.json({
           message: "PhishGuard server is running."
       });
   
   });
   
   
   /* ================================
      Start Server
      ================================ */
   
   app.listen(PORT, () => {
   
       console.log(
           `PhishGuard server running on http://localhost:${PORT}`
       );
   
   });
   /* ================================
   VirusTotal URL Check
   ================================ */

app.post("/api/check-url", async (req, res) => {

    const url = req.body.url;

    // Make sure a URL was provided
    if (!url) {
        return res.status(400).json({
            error: "No URL was provided."
        });
    }

    try {

        // Convert the URL into the format VirusTotal requires
        const urlId = Buffer
            .from(url)
            .toString("base64url");

        // Request the existing VirusTotal report
        const response = await fetch(
            `https://www.virustotal.com/api/v3/urls/${urlId}`,
            {
                method: "GET",

                headers: {
                    "x-apikey": process.env.VIRUSTOTAL_API_KEY
                }
            }
        );


        // URL has not been analyzed by VirusTotal
        if (response.status === 404) {

            return res.json({
                found: false,
                message: "No existing VirusTotal report was found."
            });

        }


        // VirusTotal returned another error
        if (!response.ok) {

            const errorText = await response.text();

            console.error(
                "VirusTotal error:",
                response.status,
                errorText
            );

            return res.status(response.status).json({
                error: "VirusTotal request failed."
            });

        }


        // Read the VirusTotal response
        const data = await response.json();

        const stats =
            data.data.attributes.last_analysis_stats;


        // Return only the information PhishGuard needs
        res.json({
            found: true,

            harmless: stats.harmless,
            malicious: stats.malicious,
            suspicious: stats.suspicious,
            undetected: stats.undetected,
            timeout: stats.timeout
        });

    } catch (error) {

        console.error(
            "VirusTotal request error:",
            error
        );

        res.status(500).json({
            error: "Unable to contact VirusTotal."
        });

    }

});