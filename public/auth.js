const AUTH_KEY = 'studyLeaveAuth';

function getAuth() {
  const raw = localStorage.getItem(AUTH_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function setAuth(payload) {
  localStorage.setItem(AUTH_KEY, JSON.stringify(payload));
}

function clearAuth() {
  localStorage.removeItem(AUTH_KEY);
}

function requireAuth() {
  const auth = getAuth();
  if (!auth) {
    window.location.href = 'login.html';
  }
}

function renderAuthStatus() {
  const auth = getAuth();
  const nameEl = document.getElementById('auth-name');
  const logoutBtn = document.getElementById('auth-logout');

  if (!nameEl || !logoutBtn) return;

  if (!auth) {
    nameEl.textContent = 'ผู้ใช้งาน';
    logoutBtn.textContent = 'เข้าสู่ระบบ';
    logoutBtn.onclick = () => {
      window.location.href = 'login.html';
    };
    return;
  }

  nameEl.textContent = auth.full_name || auth.username || 'ผู้ใช้งาน';
  logoutBtn.textContent = 'ออกจากระบบ';
  logoutBtn.onclick = () => {
    clearAuth();
    window.location.href = 'login.html';
  };
}
