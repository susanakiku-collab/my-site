const supabaseClient = window.supabase.createClient(
  window.APP_CONFIG.SUPABASE_URL,
  window.APP_CONFIG.SUPABASE_ANON_KEY
);

const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginBtn = document.getElementById("loginBtn");
const authMessage = document.getElementById("authMessage");

function setMessage(message, isError = true) {
  authMessage.textContent = message;
  authMessage.style.color = isError ? "#ff6b6b" : "#7be2ab";
}

async function handleLogin() {
  try {
    setMessage("ログイン中...", false);

    const email = loginEmail.value.trim();
    const password = loginPassword.value.trim();

    if (!email || !password) {
      setMessage("ログインIDとパスワードを入力してください。");
      return;
    }

    const { data, error } = await supabaseClient.auth.signInWithPassword({
      email,
      password
    });

    if (error) {
      setMessage("ログイン失敗: " + error.message);
      return;
    }

    setMessage("ログイン成功", false);
    console.log("login success:", data);

    window.location.href = "dashboard.html";
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

loginBtn.addEventListener("click", handleLogin);