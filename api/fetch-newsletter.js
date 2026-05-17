module.exports = async function handler(req, res) {
  const { url } = req.query;
  if (!url) return res.status(400).json({ error: "url required" });

  try {
    const upstream = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0 (compatible; Bot/1.0)" },
    });
    if (!upstream.ok) return res.status(upstream.status).json({ error: `Upstream ${upstream.status}` });
    const html = await upstream.text();
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.status(200).send(html);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
};
