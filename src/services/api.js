// Centralized API base for the XAMPP app
const BASE_API = '/study-leave-web/index.php/api';

const request = async (path, options = {}) => {
  const { method = 'GET', headers = {}, body } = options;
  const res = await fetch(`${BASE_API}${path}`, {
    method,
    headers: {
      Accept: 'application/json',
      ...headers,
    },
    body,
  });

  const text = await res.text();
  let data = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = text;
    }
  }

  if (!res.ok) {
    const message = data && data.error ? data.error : `Request failed (${res.status})`;
    const error = new Error(message);
    error.status = res.status;
    error.data = data;
    throw error;
  }

  return data;
};

const apiGet = (path) => request(path);
const apiPost = (path, payload, headers = {}) =>
  request(path, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...headers,
    },
    body: JSON.stringify(payload),
  });

const apiPostForm = (path, formData) =>
  request(path, {
    method: 'POST',
    body: formData,
  });

export { BASE_API, request, apiGet, apiPost, apiPostForm };
