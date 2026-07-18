// Health endpoint for the Smallball Site Advisor — used by the wiki Active Systems httpcheck.
// Green (200) only when the function is deployed AND the API key is configured.
module.exports = (req, res) => {
  if (process.env.ANTHROPIC_API_KEY) {
    return res.status(200).json({ ok: true, system: "smallball-advisor" });
  }
  return res.status(500).json({ ok: false, reason: "ANTHROPIC_API_KEY not set" });
};
