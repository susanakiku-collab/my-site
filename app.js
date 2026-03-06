document.addEventListener("DOMContentLoaded", () => {
  const button = document.getElementById("testButton");
  const result = document.getElementById("result");

  if (button && result) {
    button.addEventListener("click", () => {
      const now = new Date();
      const text = `JavaScript動作OK | ${now.toLocaleString("ja-JP")}`;
      result.textContent = text;
    });
  }
});