// authService
// 認証責務を dashboard.js から分離

function getCurrentUserIdSafe() {
  try {
    if (typeof currentUser !== "undefined" && currentUser?.id) return currentUser.id;
  } catch (e) {}
  return window.currentUser?.id || null;
}

async function getCurrentUserIdSafeAsync() {
  const syncId = getCurrentUserIdSafe();
  if (syncId) return syncId;
  try {
    const { data, error } = await supabaseClient.auth.getUser();
    if (error) {
      console.error(error);
      return null;
    }
    const user = data?.user || null;
    if (user) {
      window.currentUser = user;
      try {
        if (typeof setCurrentUserState === "function") setCurrentUserState(user);
        else if (typeof currentUser !== "undefined") currentUser = user;
      } catch (e) {}
      return user.id || null;
    }
  } catch (error) {
    console.error(error);
  }
  return null;
}

async function ensureAuth() {
  const { data, error } = await supabaseClient.auth.getUser();

  if (error) {
    alert("ユーザー情報の取得に失敗しました");
    window.location.href = "index.html";
    return false;
  }

  const user = data?.user || null;
  window.currentUser = user;
  try {
    if (typeof setCurrentUserState === "function") setCurrentUserState(user);
    else if (typeof currentUser !== "undefined") currentUser = user;
  } catch (e) {}

  if (!user) {
    window.location.href = "index.html";
    return false;
  }

  const { error: profileError } = await supabaseClient
    .from("profiles")
    .upsert({
      id: user.id,
      email: user.email,
      display_name: user.email,
      role: "dispatcher"
    });

  if (profileError) {
    console.error(profileError);
    alert("profiles作成エラー: " + profileError.message);
    return false;
  }

  if (typeof els !== "undefined" && els?.userEmail) {
    els.userEmail.value = user.email || "";
  }
  return true;
}

async function logout() {
  await supabaseClient.auth.signOut();
  window.currentUser = null;
  try {
    if (typeof setCurrentUserState === "function") setCurrentUserState(null);
    else if (typeof currentUser !== "undefined") currentUser = null;
  } catch (e) {}
  window.location.href = "index.html";
}
