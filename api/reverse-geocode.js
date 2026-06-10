function pickCity(address = {}) {
  return address.city || address.town || address.village || address.county || "";
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  const { lat, lon } = req.query || {};
  if (!lat || !lon) {
    return res.status(400).json({ error: "lat and lon are required" });
  }

  const url = new URL("https://nominatim.openstreetmap.org/reverse");
  url.searchParams.set("format", "jsonv2");
  url.searchParams.set("lat", String(lat));
  url.searchParams.set("lon", String(lon));
  url.searchParams.set("zoom", "10");
  url.searchParams.set("addressdetails", "1");
  url.searchParams.set("accept-language", "ja");

  try {
    const response = await fetch(url, {
      headers: {
        "User-Agent": "zairyo-mokuzai-app/1.0 (contact: example@example.com)"
      }
    });
    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({
        error: data?.error || "Nominatim reverse geocoding failed"
      });
    }

    const address = data.address || {};
    return res.status(200).json({
      prefecture: address.state || "",
      city: pickCity(address)
    });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Unexpected error" });
  }
}
