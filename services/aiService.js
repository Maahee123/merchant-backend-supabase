const axios = require('axios');
require('dotenv').config();

async function parseMessage(message) {
    try {
        const response = await axios.post(process.env.AI_API_URL, {
            input: message
        }, {
            headers: {
                "Authorization": `Bearer ${process.env.AI_API_KEY}`,
                "Content-Type": "application/json"
            }
        });

        return response.data;
    } catch (error) {
        console.error("AI Error:", error.message);
        return null;
    }
}

module.exports = { parseMessage };