const supabaseClient = window.supabase.createClient(
  window.APP_CONFIG.SUPABASE_URL,
  window.APP_CONFIG.SUPABASE_ANON_KEY
);

const loginEmail = document.getElementById("loginEmail");
const loginPassword = document.getElementById("loginPassword");
const loginBtn = document.getElementById("loginBtn");
const signupBtn = document.getElementById("signupBtn");
const authMessage = document.getElementById("authMessage");

function setMessage(message, isError = true) {
  authMessage.textContent = message;
  authMessage.style.color = isError ? "#ff6b6b" : "#7be2ab";
}

async function handleSignUp() {
  try {
    setMessage("新規登録中...", false);

    const email = loginEmail.value.trim();
    const password = loginPassword.value.trim();

    if (!email || !password) {
      setMessage("メールアドレスとパスワードを入力してください。");
      return;
    }

    const { data, error } = await supabaseClient.auth.signUp({
      email,
      password
    });

    if (error) {
      setMessage("登録失敗: " + error.message);
      return;
    }

    if (data?.user) {
      setMessage("新規登録に成功しました。確認メールを確認してください。", false);
    } else {
      setMessage("新規登録が完了しました。", false);
    }
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

async function handleLogin() {
  try {
    setMessage("ログイン中...", false);

    const email = loginEmail.value.trim();
    const password = loginPassword.value.trim();

    if (!email || !password) {
      setMessage("メールアドレスとパスワードを入力してください。");
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

    console.log("login success:", data);
    setMessage("ログイン成功", false);
  } catch (err) {
    console.error(err);
    setMessage("例外エラー: " + err.message);
  }
}

loginBtn.addEventListener("click", handleLogin);
signupBtn.addEventListener("click", handleSignUp);
