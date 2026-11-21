(() => {
  const DataBook = window.DataBook || (window.DataBook = {});
  const projectUrl = 'https://hystzongovpengbiplrt.supabase.co';
  // Publishable browser key. Access to rows and files is enforced by Supabase RLS.
  const publishableKey = 'sb_publishable_Lbpi0VU-OLD38KQMCKHAHQ_JlrV1I79';
  const tableName = 'app_data';
  const bucketName = 'data-book-files';
  const redirectTo = 'https://databook-gamma.vercel.app/order-summary.html';

  let client = null;
  let currentUser = null;
  let readyPromise = null;
  let navigate = null;
  let root = null;
  let authForm = null;
  let emailInput = null;
  let sendButton = null;
  let signOutButton = null;
  let accountLabel = null;
  let statusLabel = null;
  let list = null;
  let listStatus = null;
  let refreshId = 0;
  const listeners = new Set();
  const openers = new Map();
  const saveQueues = new Map();

  function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function showStatus(text, error = false) {
    if (!statusLabel) return;
    statusLabel.textContent = text;
    statusLabel.classList.toggle('error', error);
  }

  function showListStatus(text, error = false) {
    if (!listStatus) return;
    listStatus.textContent = text;
    listStatus.classList.toggle('error', error);
  }

  function notify(type, record) {
    const event = { type, user: currentUser, record: record || null };
    listeners.forEach(listener => {
      try { listener(event); } catch (error) { console.error('Data listener failed:', error); }
    });
  }

  function updateDashboard(count) {
    const label = document.getElementById('dashboard-data-status');
    if (!label) return;
    if (!client) label.textContent = 'Cloud storage is unavailable. Reload the page to try again.';
    else if (!currentUser) label.textContent = 'Sign in to save files and choices across devices.';
    else if (count === undefined) label.textContent = 'Connected as ' + (currentUser.email || 'your account') + '.';
    else label.textContent = count + (count === 1 ? ' saved item' : ' saved items') + ' in your Data library.';
  }

  function renderAuth() {
    if (!root) return;
    const signedIn = Boolean(currentUser);
    authForm.hidden = signedIn;
    signOutButton.hidden = !signedIn;
    accountLabel.textContent = signedIn
      ? 'Connected as ' + (currentUser.email || 'your account') + '. Your saved data is private to this account.'
      : 'Sign in with your email to save and reopen your files and settings.';
    updateDashboard();
  }

  function formatDate(value) {
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? '' : date.toLocaleString();
  }

  function formatSize(bytes) {
    if (!Number.isFinite(Number(bytes))) return '';
    const size = Number(bytes);
    if (size < 1024) return size + ' B';
    if (size < 1024 * 1024) return (size / 1024).toFixed(1) + ' KB';
    return (size / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function renderRecords(records) {
    list.replaceChildren();
    if (!records.length) {
      list.append(element('li', 'data-empty', 'Nothing saved yet. Upload an order file and its settings will appear here.'));
      return;
    }
    records.forEach(record => {
      const item = element('li', 'data-item');
      const details = element('div');
      details.append(element('strong', '', record.title || record.record_key));
      const detailsText = [record.module, formatDate(record.updated_at), record.file_size !== null ? formatSize(record.file_size) : ''].filter(Boolean).join(' · ');
      details.append(element('p', 'data-item-meta', detailsText));
      const actions = element('div', 'data-actions');
      if (openers.has(record.module)) {
        const openButton = element('button', 'data-button', 'Open');
        openButton.type = 'button';
        openButton.addEventListener('click', async () => {
          openButton.disabled = true;
          showListStatus('Opening ' + (record.title || 'saved item') + '...');
          try {
            await openers.get(record.module)(record);
            if (navigate) navigate(record.module);
            showListStatus('Opened ' + (record.title || 'saved item') + '.');
          } catch (error) {
            showListStatus(error.message || 'Could not open this saved item.', true);
          } finally {
            openButton.disabled = false;
          }
        });
        actions.append(openButton);
      }
      if (record.file_path) {
        const downloadButton = element('button', 'data-button', 'Download');
        downloadButton.type = 'button';
        downloadButton.addEventListener('click', async () => {
          downloadButton.disabled = true;
          showListStatus('Downloading ' + (record.file_name || record.title) + '...');
          try {
            const blob = await downloadFile(record);
            const url = URL.createObjectURL(blob);
            const anchor = document.createElement('a');
            anchor.href = url;
            anchor.download = record.file_name || record.title || 'data-file';
            document.body.append(anchor);
            anchor.click();
            anchor.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
            showListStatus('Download started.');
          } catch (error) {
            showListStatus(error.message || 'Could not download the file.', true);
          } finally {
            downloadButton.disabled = false;
          }
        });
        actions.append(downloadButton);
      }
      const deleteButton = element('button', 'data-button', 'Delete');
      deleteButton.type = 'button';
      deleteButton.addEventListener('click', async () => {
        if (!window.confirm('Delete “' + (record.title || 'this item') + '” from Data?')) return;
        deleteButton.disabled = true;
        showListStatus('Deleting ' + (record.title || 'saved item') + '...');
        try {
          await deleteRecord(record);
          showListStatus('Saved item deleted.');
        } catch (error) {
          showListStatus(error.message || 'Could not delete this item.', true);
          await refreshRecords();
        } finally {
          deleteButton.disabled = false;
        }
      });
      actions.append(deleteButton);
      item.append(details, actions);
      list.append(item);
    });
  }

  async function refreshRecords() {
    if (!list) return;
    const request = ++refreshId;
    if (!currentUser) {
      list.replaceChildren(element('li', 'data-empty', 'Sign in to view your saved data.'));
      showListStatus('');
      updateDashboard();
      return;
    }
    showListStatus('Loading saved data...');
    try {
      const records = await listRecords();
      if (request !== refreshId) return;
      renderRecords(records);
      showListStatus('');
      updateDashboard(records.length);
    } catch (error) {
      if (request !== refreshId) return;
      list.replaceChildren(element('li', 'data-empty', 'Saved data could not be loaded.'));
      showListStatus(error.message || 'Could not load saved data.', true);
      updateDashboard();
    }
  }

  function mount() {
    root = document.getElementById('data-module-root');
    if (!root) return;
    const accountPanel = element('section', 'data-panel');
    accountPanel.append(element('h2', '', 'Cloud connection'));
    accountLabel = element('p');
    accountPanel.append(accountLabel);
    authForm = element('form', 'data-form');
    const emailLabel = element('label', '', 'Email address');
    emailInput = element('input', 'data-input');
    emailInput.type = 'email';
    emailInput.autocomplete = 'email';
    emailInput.required = true;
    emailInput.placeholder = 'you@example.com';
    emailLabel.append(emailInput);
    sendButton = element('button', 'data-button', 'Send sign-in link');
    sendButton.type = 'submit';
    authForm.append(emailLabel, sendButton);
    authForm.addEventListener('submit', async event => {
      event.preventDefault();
      if (!client) return;
      sendButton.disabled = true;
      showStatus('Sending sign-in link...');
      const email = emailInput.value.trim();
      try {
        const { error } = await client.auth.signInWithOtp({
          email,
          options: { emailRedirectTo: redirectTo }
        });
        if (error) throw error;
        showStatus('Check your email and open the sign-in link in this browser.');
      } catch (error) {
        showStatus(error.message || 'Could not send a sign-in link.', true);
      } finally {
        sendButton.disabled = false;
      }
    });
    accountPanel.append(authForm);
    signOutButton = element('button', 'data-button', 'Sign out');
    signOutButton.type = 'button';
    signOutButton.hidden = true;
    signOutButton.addEventListener('click', async () => {
      signOutButton.disabled = true;
      try {
        const { error } = await client.auth.signOut();
        if (error) throw error;
        showStatus('Signed out.');
      } catch (error) {
        showStatus(error.message || 'Could not sign out.', true);
      } finally {
        signOutButton.disabled = false;
      }
    });
    accountPanel.append(signOutButton);
    statusLabel = element('p', 'data-status');
    statusLabel.setAttribute('role', 'status');
    statusLabel.setAttribute('aria-live', 'polite');
    accountPanel.append(statusLabel);

    const libraryPanel = element('section', 'data-panel');
    libraryPanel.append(element('h2', '', 'Saved data'));
    libraryPanel.append(element('p', '', 'Files and settings from every module are kept here.'));
    listStatus = element('p', 'data-status');
    listStatus.setAttribute('role', 'status');
    listStatus.setAttribute('aria-live', 'polite');
    libraryPanel.append(listStatus);
    list = element('ul', 'data-list');
    libraryPanel.append(list);
    root.replaceChildren(accountPanel, libraryPanel);
    renderAuth();
    refreshRecords();
  }

  async function requireUser() {
    await ready();
    if (!client) throw new Error('Cloud storage is unavailable. Reload the page to try again.');
    if (!currentUser) throw new Error('Sign in on the Data page before saving or reading cloud data.');
    return currentUser;
  }

  function validateIdentity(module, key) {
    if (typeof module !== 'string' || !/^[a-z0-9][a-z0-9._-]{0,79}$/i.test(module)) {
      throw new Error('Invalid data module name.');
    }
    if (typeof key !== 'string' || !key || key.length > 256) {
      throw new Error('Invalid data record key.');
    }
  }

  async function getRecord(module, key) {
    validateIdentity(module, key);
    const owner = await requireUser();
    const { data, error } = await client.from(tableName).select('*')
      .eq('user_id', owner.id).eq('module', module).eq('record_key', key).maybeSingle();
    if (error) throw error;
    return data || null;
  }

  async function listRecords(module) {
    if (module !== undefined) validateIdentity(module, 'list');
    const owner = await requireUser();
    const records = [];
    const pageSize = 500;
    for (let offset = 0; ; offset += pageSize) {
      let query = client.from(tableName).select('*').eq('user_id', owner.id)
        .order('updated_at', { ascending: false }).range(offset, offset + pageSize - 1);
      if (module !== undefined) query = query.eq('module', module);
      const { data, error } = await query;
      if (error) throw error;
      records.push(...data);
      if (data.length < pageSize) break;
    }
    return records;
  }

  async function saveRecordNow({ module, key, title, payload = {}, file }) {
    validateIdentity(module, key);
    if (typeof title !== 'string' || !title.trim()) throw new Error('A saved item needs a title.');
    const owner = await requireUser();
    const existing = await getRecord(module, key);
    let uploadedPath = null;
    let filePath = existing?.file_path || null;
    let fileName = existing?.file_name || null;
    let fileSize = existing?.file_size ?? null;
    try {
      if (file && !filePath) {
        if (!(file instanceof Blob)) throw new Error('The file must be a browser File or Blob.');
        fileName = file.name || title;
        fileSize = file.size;
        const extension = /\.([a-z0-9]{1,12})$/i.exec(fileName)?.[1]?.toLowerCase() || 'bin';
        const fileId = crypto.randomUUID();
        uploadedPath = owner.id + '/' + module + '/' + fileId + '.' + extension;
        const { error } = await client.storage.from(bucketName).upload(uploadedPath, file, {
          contentType: file.type || 'application/octet-stream', upsert: false
        });
        if (error) throw error;
        filePath = uploadedPath;
      }
      const values = {
        title: title.trim(), payload, file_path: filePath, file_name: fileName,
        file_size: fileSize, updated_at: new Date().toISOString()
      };
      const query = existing
        ? client.from(tableName).update(values).eq('id', existing.id).eq('user_id', owner.id)
        : client.from(tableName).insert({ ...values, user_id: owner.id, module, record_key: key });
      const { data, error } = await query.select('*').single();
      if (error) throw error;
      notify('records', data);
      refreshRecords();
      return data;
    } catch (error) {
      if (uploadedPath) {
        const { error: cleanupError } = await client.storage.from(bucketName).remove([uploadedPath]);
        if (cleanupError) console.error('Could not clean up uploaded file:', cleanupError);
      }
      throw error;
    }
  }

  function saveRecord(record) {
    if (!record || typeof record !== 'object') return Promise.reject(new Error('Provide a data record to save.'));
    const queueKey = String(record.module) + '\u0000' + String(record.key);
    const prior = saveQueues.get(queueKey) || Promise.resolve();
    const task = prior.catch(() => {}).then(() => saveRecordNow(record));
    saveQueues.set(queueKey, task);
    task.finally(() => { if (saveQueues.get(queueKey) === task) saveQueues.delete(queueKey); }).catch(() => {});
    return task;
  }

  async function downloadFile(record) {
    if (!record?.file_path) throw new Error('This saved item has no file.');
    const owner = await requireUser();
    if (record.user_id && record.user_id !== owner.id) throw new Error('You cannot open another account’s file.');
    const { data, error } = await client.storage.from(bucketName).download(record.file_path);
    if (error) throw error;
    return data;
  }

  async function deleteRecord(record) {
    if (!record?.id) throw new Error('Choose a saved item to delete.');
    const owner = await requireUser();
    const queuedSave = saveQueues.get(record.module + '\u0000' + record.record_key);
    if (queuedSave) await queuedSave.catch(() => {});
    const { data, error } = await client.from(tableName).delete().eq('id', record.id)
      .eq('user_id', owner.id).select('id').maybeSingle();
    if (error) throw error;
    if (!data) throw new Error('This saved item no longer exists.');
    notify('deleted', record);
    refreshRecords();
    if (record.file_path) {
      const { error: storageError } = await client.storage.from(bucketName).remove([record.file_path]);
      if (storageError) throw new Error('Saved item deleted, but its stored file could not be cleaned up: ' + storageError.message);
    }
  }

  function registerOpener(module, callback) {
    validateIdentity(module, 'open');
    if (typeof callback !== 'function') throw new Error('An opener must be a function.');
    openers.set(module, callback);
    refreshRecords();
    return () => { if (openers.get(module) === callback) openers.delete(module); };
  }

  function onChange(callback) {
    if (typeof callback !== 'function') throw new Error('A data listener must be a function.');
    listeners.add(callback);
    return () => listeners.delete(callback);
  }

  function ready() {
    return readyPromise || Promise.reject(new Error('The Data module has not started.'));
  }

  function init(showView) {
    if (readyPromise) return readyPromise;
    navigate = showView;
    mount();
    if (!window.supabase?.createClient) {
      showStatus('Supabase could not load. Check your connection and reload the page.', true);
      updateDashboard();
      readyPromise = Promise.reject(new Error('Supabase library did not load.'));
      readyPromise.catch(() => {});
      return readyPromise;
    }
    client = window.supabase.createClient(projectUrl, publishableKey, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true, flowType: 'implicit' }
    });
    client.auth.onAuthStateChange((_event, session) => {
      currentUser = session?.user || null;
      setTimeout(() => {
        renderAuth();
        notify('auth');
        refreshRecords();
      }, 0);
    });
    readyPromise = (async () => {
      const { data, error } = await client.auth.getSession();
      if (error) throw error;
      currentUser = data.session?.user || null;
      renderAuth();
      notify('auth');
      return currentUser;
    })().catch(error => {
      showStatus(error.message || 'Could not connect to Supabase.', true);
      throw error;
    });
    readyPromise.then(() => refreshRecords(), () => {});
    readyPromise.catch(() => {});
    return readyPromise;
  }

  DataBook.data = {
    init, ready, user: () => currentUser, onChange, getRecord, listRecords,
    saveRecord, downloadFile, deleteRecord, registerOpener
  };
})();
