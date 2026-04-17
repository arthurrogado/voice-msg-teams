const STORAGE_KEY_PREFERRED_MICROPHONE_ID = 'preferredMicrophoneId';
const STORAGE_KEY_NOISE_SUPPRESSION_ENABLED = 'noiseSuppressionEnabled';
const SYSTEM_DEFAULT_MICROPHONE_ID = 'default';
const DEFAULT_NOISE_SUPPRESSION_ENABLED = true;

const selectElement = document.getElementById('microphone-select');
const refreshButton = document.getElementById('refresh-button');
const noiseSuppressionToggle = document.getElementById('noise-suppression-toggle');
const statusElement = document.getElementById('status');

initialize().catch((error) => {
  setStatus(error?.message || message('popupStatusLoadFailed'));
});

async function initialize() {
  applyI18n();
  attachEvents();
  await loadNoiseSuppressionPreference();
  await loadMicrophones();
}

function applyI18n() {
  document.title = message('actionTitle');

  for (const element of document.querySelectorAll('[data-i18n]')) {
    const key = element.getAttribute('data-i18n');
    element.textContent = message(key);
  }

  for (const element of document.querySelectorAll('[data-i18n-aria-label]')) {
    const key = element.getAttribute('data-i18n-aria-label');
    element.setAttribute('aria-label', message(key));
    element.setAttribute('title', message(key));
  }
}

function attachEvents() {
  refreshButton.addEventListener('click', () => {
    loadMicrophones().catch((error) => {
      setStatus(error?.message || message('popupStatusLoadFailed'));
    });
  });

  selectElement.addEventListener('change', async () => {
    await savePreferredMicrophoneId(selectElement.value);
    setStatus(message('popupStatusSaved'));
  });

  noiseSuppressionToggle.addEventListener('change', async () => {
    await saveNoiseSuppressionEnabled(noiseSuppressionToggle.checked);
    setStatus(message('popupStatusNoiseSuppressionSaved'));
  });
}

async function loadNoiseSuppressionPreference() {
  noiseSuppressionToggle.checked = await getNoiseSuppressionEnabled();
}

async function loadMicrophones() {
  setLoading(true);
  setStatus(message('popupStatusLoading'));

  const storedPreference = await getPreferredMicrophoneId();
  const tab = await findTeamsTab();

  if (!tab?.id) {
    populateOptions([], storedPreference);
    setStatus(message('popupStatusOpenTeams'));
    setLoading(false);
    return;
  }

  const response = await sendMessageToTab(tab.id, { type: 'teamsVoiceMessage.listMicrophones' });

  if (!response?.ok) {
    populateOptions([], storedPreference);
    setStatus(response?.error || message('popupStatusLoadFailed'));
    setLoading(false);
    return;
  }

  const devices = Array.isArray(response.devices) ? response.devices : [];
  const availableDeviceIds = new Set(devices.map((device) => device.deviceId));
  let preferredMicrophoneId = storedPreference;

  if (preferredMicrophoneId !== SYSTEM_DEFAULT_MICROPHONE_ID && !availableDeviceIds.has(preferredMicrophoneId)) {
    preferredMicrophoneId = SYSTEM_DEFAULT_MICROPHONE_ID;
    await savePreferredMicrophoneId(SYSTEM_DEFAULT_MICROPHONE_ID);
    setStatus(message('popupStatusSelectedMissing'));
  } else if (devices.length === 0) {
    setStatus(message('microphoneNotFound'));
  } else if (!response.labelsAvailable) {
    setStatus(message('popupStatusLabelsHidden'));
  } else {
    setStatus(message('popupStatusReady'));
  }

  populateOptions(devices, preferredMicrophoneId);
  setLoading(false);
}

function populateOptions(devices, preferredMicrophoneId) {
  selectElement.textContent = '';

  const defaultOption = document.createElement('option');
  defaultOption.value = SYSTEM_DEFAULT_MICROPHONE_ID;
  defaultOption.textContent = message('popupSystemDefaultOption');
  selectElement.appendChild(defaultOption);

  for (const device of devices) {
    const option = document.createElement('option');
    option.value = device.deviceId;
    option.textContent = device.label;
    selectElement.appendChild(option);
  }

  selectElement.value = preferredMicrophoneId || SYSTEM_DEFAULT_MICROPHONE_ID;
}

function setLoading(isLoading) {
  selectElement.disabled = isLoading;
  refreshButton.disabled = isLoading;
  noiseSuppressionToggle.disabled = isLoading;
}

function setStatus(text) {
  statusElement.textContent = text || '';
}

function message(key) {
  return chrome.i18n?.getMessage?.(key) || key;
}

function getPreferredMicrophoneId() {
  return new Promise((resolve) => {
    chrome.storage.sync.get({ [STORAGE_KEY_PREFERRED_MICROPHONE_ID]: SYSTEM_DEFAULT_MICROPHONE_ID }, (items) => {
      if (chrome.runtime.lastError) {
        resolve(SYSTEM_DEFAULT_MICROPHONE_ID);
        return;
      }

      const preferredMicrophoneId = items?.[STORAGE_KEY_PREFERRED_MICROPHONE_ID];
      resolve(typeof preferredMicrophoneId === 'string' && preferredMicrophoneId ? preferredMicrophoneId : SYSTEM_DEFAULT_MICROPHONE_ID);
    });
  });
}

function getNoiseSuppressionEnabled() {
  return new Promise((resolve) => {
    chrome.storage.sync.get({ [STORAGE_KEY_NOISE_SUPPRESSION_ENABLED]: DEFAULT_NOISE_SUPPRESSION_ENABLED }, (items) => {
      if (chrome.runtime.lastError) {
        resolve(DEFAULT_NOISE_SUPPRESSION_ENABLED);
        return;
      }

      resolve(items?.[STORAGE_KEY_NOISE_SUPPRESSION_ENABLED] !== false);
    });
  });
}

function savePreferredMicrophoneId(preferredMicrophoneId) {
  return new Promise((resolve) => {
    chrome.storage.sync.set({ [STORAGE_KEY_PREFERRED_MICROPHONE_ID]: preferredMicrophoneId }, () => {
      resolve();
    });
  });
}

function saveNoiseSuppressionEnabled(isEnabled) {
  return new Promise((resolve) => {
    chrome.storage.sync.set({ [STORAGE_KEY_NOISE_SUPPRESSION_ENABLED]: Boolean(isEnabled) }, () => {
      resolve();
    });
  });
}

function findTeamsTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      resolve(Array.isArray(tabs) && tabs.length > 0 ? tabs[0] : null);
    });
  });
}

function sendMessageToTab(tabId, payload) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, payload, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: message('popupStatusOpenTeams') });
        return;
      }

      resolve(response);
    });
  });
}