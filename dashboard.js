async function geocodeAddress(address) {

  const normalized = String(address || "")
    .replace(/〒\s*\d{3}-?\d{4}/g, "")
    .replace(/^日本[、,\s]*/g, "")
    .trim();

  if (!normalized) {
    throw new Error("住所が空です");
  }

  const geocoder = new google.maps.Geocoder();

  return new Promise((resolve, reject) => {

    geocoder.geocode(
      {
        address: normalized,
        region: "jp"
      },
      (results, status) => {

        if (status !== "OK") {
          reject(new Error("Geocoding失敗: " + status));
          return;
        }

        const loc = results[0].geometry.location;

        resolve({
          lat: loc.lat(),
          lng: loc.lng()
        });

      }
    );

  });

}
