(() => {
  const INSTANCE_KEY = '__teamsVoiceMessageInstance';
  const ROOT_ID = 'teams-voice-message-root';
  const BUTTON_ID = 'teams-voice-message-button';
  const BUBBLE_ID = 'teams-voice-message-bubble';
  const previousInstance = window[INSTANCE_KEY];

  if (previousInstance && typeof previousInstance.dispose === 'function') {
    try {
      previousInstance.dispose('reinitialize');
    } catch (_error) {
      document.getElementById(ROOT_ID)?.remove();
      document.getElementById(BUBBLE_ID)?.remove();
    }
  }

  window.__teamsVoiceMessageInitialized = true;

  const STORAGE_KEY_PREFERRED_MICROPHONE_ID = 'preferredMicrophoneId';
  const STORAGE_KEY_NOISE_SUPPRESSION_ENABLED = 'noiseSuppressionEnabled';
  const SYSTEM_DEFAULT_MICROPHONE_ID = 'default';
  const DEFAULT_NOISE_SUPPRESSION_ENABLED = true;
  const RECORDING_LIMIT_MS = 5 * 60 * 1000;
  const NATIVE_FILE_PICKER_SELECTOR = '[data-tid="sendMessageCommands-FilePicker"], [data-tid="newMessageCommands-FilePicker"]';
  const SEND_BUTTON_SELECTOR = '[data-tid="sendMessageCommands-send"], [data-tid="newMessageCommands-send"]';
  const MIME_TYPE_CANDIDATES = [
    'audio/webm;codecs=opus',
    'audio/webm',
    'audio/ogg;codecs=opus'
  ];
  const EDITOR_SELECTORS = [
    '[contenteditable="true"][role="textbox"]',
    '[contenteditable="true"][aria-label]',
    '[contenteditable="true"][data-lexical-editor="true"]',
    '[contenteditable="true"]',
    'div[role="textbox"]',
    'textarea'
  ];

  const state = {
    isBusy: false,
    isRecording: false,
    isContextInvalidated: false,
    mediaRecorder: null,
    mediaStream: null,
    chunks: [],
    mimeType: '',
    startedAt: 0,
    timerId: null,
    autoStopId: null,
    messageTimerId: null,
    messageText: '',
    messageTone: 'info',
    pendingAttachmentName: '',
    observer: null,
    refreshRaf: 0,
    runtimeMessageListener: null
  };

  window[INSTANCE_KEY] = {
    dispose: invalidateExtensionContext
  };

  function message(key) {
    if (state.isContextInvalidated) {
      return key;
    }

    try {
      return chrome.i18n?.getMessage?.(key) || key;
    } catch (error) {
      if (isExtensionContextInvalidated(error)) {
        invalidateExtensionContext();
        return key;
      }

      throw error;
    }
  }

  function withDuration(key, duration) {
    return message(key) + ' ' + duration;
  }

  function bootstrap() {
    if (state.isContextInvalidated) {
      return;
    }

    installRuntimeListeners();
    installObservers();
    scheduleRefresh();
  }

  function installRuntimeListeners() {
    if (state.isContextInvalidated || !chrome.runtime?.onMessage || window.__teamsVoiceMessageRuntimeListenersInstalled) {
      return;
    }

    window.__teamsVoiceMessageRuntimeListenersInstalled = true;
    state.runtimeMessageListener = (request, _sender, sendResponse) => {
      if (request?.type !== 'teamsVoiceMessage.listMicrophones') {
        return false;
      }

      listAvailableMicrophones()
        .then((result) => {
          sendResponse({ ok: true, ...result });
        })
        .catch((error) => {
          sendResponse({
            ok: false,
            error: normalizeError(error)
          });
        });

      return true;
    };

    chrome.runtime.onMessage.addListener(state.runtimeMessageListener);
  }

  function installObservers() {
    if (state.observer || !(document.documentElement instanceof HTMLElement)) {
      return;
    }

    state.observer = new MutationObserver(() => {
      scheduleRefresh();
    });

    state.observer.observe(document.body || document.documentElement, {
      childList: true,
      subtree: true
    });

    window.addEventListener('resize', scheduleRefresh, { passive: true });
    window.addEventListener('focusin', scheduleRefresh, true);
  }

  function scheduleRefresh() {
    if (state.isContextInvalidated || state.refreshRaf) {
      return;
    }

    state.refreshRaf = window.requestAnimationFrame(() => {
      state.refreshRaf = 0;

      try {
        refreshMount();
      } catch (error) {
        if (isExtensionContextInvalidated(error)) {
          invalidateExtensionContext();
          return;
        }

        throw error;
      }
    });
  }

  function refreshMount() {
    if (state.isContextInvalidated) {
      return;
    }

    const context = findComposerContext();

    if (!context) {
      state.pendingAttachmentName = '';

      if (!state.isRecording && !state.isBusy) {
        removeRoot();
      }

      return;
    }

    if (state.pendingAttachmentName && !detectAttachment(context.composeRoot, state.pendingAttachmentName).ok) {
      state.pendingAttachmentName = '';
    }

    const root = ensureRoot();

    if (!(root instanceof HTMLElement)) {
      return;
    }

    if (!(context.mountTarget instanceof HTMLElement) || context.mountTarget.parentElement !== context.toolbar) {
      if (root.parentElement !== context.toolbar) {
        context.toolbar.appendChild(root);
      }
    } else if (root.parentElement !== context.toolbar || root.nextSibling !== context.mountTarget) {
      context.toolbar.insertBefore(root, context.mountTarget);
    }

    syncUi();
  }

  function findComposerContext() {
    const editor = findActiveEditor();
    const anchorButton = document.querySelector(NATIVE_FILE_PICKER_SELECTOR) || document.querySelector(SEND_BUTTON_SELECTOR);
    const toolbar = anchorButton instanceof HTMLElement ? anchorButton.closest('[role="toolbar"]') : null;

    if (!(editor instanceof HTMLElement) || !(toolbar instanceof HTMLElement) || !(anchorButton instanceof HTMLElement)) {
      return null;
    }

    return {
      editor,
      composeRoot: findComposeRoot(editor),
      toolbar,
      anchorButton,
      mountTarget: findToolbarMountTarget(toolbar, anchorButton)
    };
  }

  function findToolbarMountTarget(toolbar, anchorButton) {
    if (!(toolbar instanceof HTMLElement) || !(anchorButton instanceof HTMLElement)) {
      return null;
    }

    let current = anchorButton;

    while (current.parentElement && current.parentElement !== toolbar) {
      current = current.parentElement;
    }

    return current.parentElement === toolbar ? current : null;
  }

  function ensureRoot() {
    if (!(document.body instanceof HTMLBodyElement)) {
      return null;
    }

    let root = document.getElementById(ROOT_ID);

    if (root) {
      return root;
    }

    root = document.createElement('div');
    root.id = ROOT_ID;
    root.className = 'tvm-toolbar-item';
    root.style.setProperty('--tvm-icon-url', 'url("' + runtimeUrl('mic.svg') + '")');

    const button = document.createElement('button');
    button.id = BUTTON_ID;
    button.className = 'tvm-button';
    button.type = 'button';
    button.setAttribute('aria-label', message('recordButtonLabel'));
    button.setAttribute('title', message('recordButtonLabel'));
    button.addEventListener('click', onToggleRecording);

    root.appendChild(button);

    return root;
  }

  function ensureBubble() {
    if (!(document.body instanceof HTMLBodyElement)) {
      return null;
    }

    let bubble = document.getElementById(BUBBLE_ID);

    if (bubble) {
      return bubble;
    }

    bubble = document.createElement('div');
    bubble.id = BUBBLE_ID;
    bubble.className = 'tvm-bubble';
    bubble.hidden = true;
    bubble.setAttribute('aria-live', 'polite');
    document.body.appendChild(bubble);
    return bubble;
  }

  function removeRoot() {
    document.getElementById(ROOT_ID)?.remove();
  }

  async function onToggleRecording() {
    if (state.isContextInvalidated || state.isBusy) {
      return;
    }

    state.isBusy = true;
    syncUi();

    try {
      if (state.isRecording) {
        await stopRecordingAndAttach();
      } else {
        await startRecording();
      }
    } catch (error) {
      cleanupRecorder();
      showMessage(normalizeError(error), 'error', 4200);
    } finally {
      state.isBusy = false;
      syncUi();
      scheduleRefresh();
    }
  }

  async function startRecording() {
    if (!navigator.mediaDevices || typeof navigator.mediaDevices.getUserMedia !== 'function') {
      throw new Error(message('browserUnsupportedMicrophoneCapture'));
    }

    const context = findComposerContext();

    if (!context) {
      throw new Error(message('composerNotFound'));
    }

    if (hasPendingAttachment(context.composeRoot)) {
      throw new Error(message('sendOrRemovePendingAttachment'));
    }

    const stream = await requestAudioStream();

    const mimeType = chooseMimeType();
    const recorderOptions = mimeType ? { mimeType } : undefined;
    let recorder;

    try {
      recorder = recorderOptions ? new MediaRecorder(stream, recorderOptions) : new MediaRecorder(stream);
    } catch (_error) {
      stopStream(stream);
      throw new Error(message('startRecordingFailed'));
    }

    state.mediaStream = stream;
    state.mediaRecorder = recorder;
    state.mimeType = recorder.mimeType || mimeType || 'audio/webm';
    state.chunks = [];
    state.startedAt = Date.now();
    state.isRecording = true;

    recorder.addEventListener('dataavailable', handleDataAvailable);
    recorder.addEventListener('error', handleRecorderError, { once: true });
    recorder.start(300);

    clearTimeout(state.autoStopId);
    state.autoStopId = window.setTimeout(() => {
      if (state.isRecording) {
        onToggleRecording();
        showMessage(message('recordingStoppedLimit'), 'warning', 2600);
      }
    }, RECORDING_LIMIT_MS);

    startTimer();
    showPersistentMessage(withDuration('recordingInProgressPrefix', '00:00'), 'recording');
  }

  async function stopRecordingAndAttach() {
    const recorder = state.mediaRecorder;

    if (!recorder) {
      throw new Error(message('noActiveRecording'));
    }

    showPersistentMessage(message('convertingToWav'), 'info');
    await stopRecorder(recorder);

    const sourceBlob = new Blob(state.chunks, {
      type: state.mimeType || recorder.mimeType || 'audio/webm'
    });

    if (!sourceBlob.size) {
      cleanupRecorder();
      throw new Error(message('emptyRecording'));
    }

    const file = await buildAudioFile(sourceBlob);

    showPersistentMessage(message('attachingAudio'), 'info');
    await attachFileToTeams(file);
    state.pendingAttachmentName = file.name;

    cleanupRecorder();
    clearMessage();
  }

  function stopRecorder(recorder) {
    return new Promise((resolve, reject) => {
      if (recorder.state === 'inactive') {
        resolve();
        return;
      }

      const handleStop = () => {
        recorder.removeEventListener('stop', handleStop);
        recorder.removeEventListener('error', handleError);
        resolve();
      };

      const handleError = () => {
        recorder.removeEventListener('stop', handleStop);
        recorder.removeEventListener('error', handleError);
        reject(new Error(message('stopRecordingFailed')));
      };

      recorder.addEventListener('stop', handleStop, { once: true });
      recorder.addEventListener('error', handleError, { once: true });
      recorder.stop();
    });
  }

  function handleDataAvailable(event) {
    if (event.data && event.data.size > 0) {
      state.chunks.push(event.data);
    }
  }

  function handleRecorderError() {
    cleanupRecorder();
    showMessage(message('browserReportedRecordingError'), 'error', 4200);
  }

  async function buildAudioFile(blob) {
    const wavBlob = await convertBlobToWav(blob);
    const timestamp = new Date().toISOString().replace(/[.:]/g, '-');
    const fileName = message('voiceMessageFilePrefix') + '_' + timestamp + '.wav';

    return new File([wavBlob], fileName, {
      type: 'audio/wav',
      lastModified: Date.now()
    });
  }

  async function convertBlobToWav(blob) {
    const sourceBuffer = await blob.arrayBuffer();
    const audioContext = new AudioContext();

    try {
      const decoded = await audioContext.decodeAudioData(sourceBuffer.slice(0));
      const wavBuffer = encodeAudioBufferToWav(decoded);
      return new Blob([wavBuffer], { type: 'audio/wav' });
    } catch (_error) {
      throw new Error(message('convertToWavFailed'));
    } finally {
      await audioContext.close().catch(() => undefined);
    }
  }

  function encodeAudioBufferToWav(audioBuffer) {
    const channelCount = audioBuffer.numberOfChannels;
    const sampleRate = audioBuffer.sampleRate;
    const monoSamples = mixDownToMono(audioBuffer, channelCount);
    const bytesPerSample = 2;
    const blockAlign = bytesPerSample;
    const byteRate = sampleRate * blockAlign;
    const wavBuffer = new ArrayBuffer(44 + monoSamples.length * bytesPerSample);
    const view = new DataView(wavBuffer);

    writeAscii(view, 0, 'RIFF');
    view.setUint32(4, 36 + monoSamples.length * bytesPerSample, true);
    writeAscii(view, 8, 'WAVE');
    writeAscii(view, 12, 'fmt ');
    view.setUint32(16, 16, true);
    view.setUint16(20, 1, true);
    view.setUint16(22, 1, true);
    view.setUint32(24, sampleRate, true);
    view.setUint32(28, byteRate, true);
    view.setUint16(32, blockAlign, true);
    view.setUint16(34, 16, true);
    writeAscii(view, 36, 'data');
    view.setUint32(40, monoSamples.length * bytesPerSample, true);

    let offset = 44;

    for (let index = 0; index < monoSamples.length; index += 1) {
      const sample = Math.max(-1, Math.min(1, monoSamples[index]));
      view.setInt16(offset, sample < 0 ? sample * 0x8000 : sample * 0x7fff, true);
      offset += 2;
    }

    return wavBuffer;
  }

  function mixDownToMono(audioBuffer, channelCount) {
    if (channelCount === 1) {
      return new Float32Array(audioBuffer.getChannelData(0));
    }

    const frameCount = audioBuffer.length;
    const mono = new Float32Array(frameCount);

    for (let channelIndex = 0; channelIndex < channelCount; channelIndex += 1) {
      const channelData = audioBuffer.getChannelData(channelIndex);

      for (let frameIndex = 0; frameIndex < frameCount; frameIndex += 1) {
        mono[frameIndex] += channelData[frameIndex] / channelCount;
      }
    }

    return mono;
  }

  function writeAscii(view, offset, text) {
    for (let index = 0; index < text.length; index += 1) {
      view.setUint8(offset + index, text.charCodeAt(index));
    }
  }

  async function attachFileToTeams(file) {
    const context = findComposerContext();

    if (!context) {
      throw new Error(message('composerNotFound'));
    }

    focusElement(context.editor);

    const nativePickerResult = await attachViaNativePicker(file, context);
    if (nativePickerResult.ok) {
      return nativePickerResult;
    }

    const directInputResult = await attachViaExistingInput(file, context.composeRoot || context.toolbar);
    if (directInputResult.ok) {
      return directInputResult;
    }

    throw new Error(nativePickerResult.message || directInputResult.message || message('teamsDidNotConfirmAttachment'));
  }

  async function attachViaNativePicker(file, context) {
    const input = await ensureNativeFilePickerInput(context);

    if (!(input instanceof HTMLInputElement)) {
      return {
        ok: false,
        message: message('teamsCouldNotOpenFilePicker')
      };
    }

    const applied = applyFileToInput(input, file);

    if (!applied) {
      return {
        ok: false,
        message: message('teamsFilePickerRejectedAudio')
      };
    }

    const verification = await waitForAttachment(context.composeRoot, file.name);

    if (verification.ok) {
      return {
        ok: true,
        method: 'picker-nativo',
        verification
      };
    }

    return {
      ok: false,
      message: message('teamsDidNotConfirmWavUpload')
    };
  }

  async function attachViaExistingInput(file, scope) {
    const composeScope = scope instanceof HTMLElement ? scope : document.body;
    const inputs = collectFileInputs([composeScope, composeScope.parentElement, document.body]);

    for (const input of inputs) {
      if (!(input instanceof HTMLInputElement)) {
        continue;
      }

      const applied = applyFileToInput(input, file);
      if (!applied) {
        continue;
      }

      const verification = await waitForAttachment(composeScope, file.name);

      if (verification.ok) {
        return {
          ok: true,
          method: 'input-direto',
          verification
        };
      }
    }

    return {
      ok: false,
      message: 'Nenhum input de arquivo reutilizavel foi encontrado.'
    };
  }

  function collectFileInputs(scopes) {
    const inputs = [];
    const seen = new Set();

    for (const scope of uniqueElements(scopes)) {
      for (const input of scope.querySelectorAll('input[type="file"]')) {
        if (!(input instanceof HTMLInputElement) || seen.has(input)) {
          continue;
        }

        seen.add(input);
        inputs.push(input);
      }
    }

    return inputs;
  }

  function applyFileToInput(input, file) {
    try {
      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(file);
      input.files = dataTransfer.files;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    } catch (_error) {
      return false;
    }
  }

  async function ensureNativeFilePickerInput(context) {
    const knownInputs = collectFileInputs([context.composeRoot, context.toolbar, document.body]);
    const knownSet = new Set(knownInputs);
    const existingInput = chooseBestFileInput(knownInputs, context);

    if (existingInput instanceof HTMLInputElement) {
      return existingInput;
    }

    if (!(context.anchorButton instanceof HTMLElement)) {
      return null;
    }

    dispatchPointerClick(context.anchorButton);

    for (const delay of [120, 260, 520, 900]) {
      await wait(delay);
      const inputs = collectFileInputs([context.composeRoot, context.toolbar, document.body]);
      const freshInput = inputs.find((input) => !knownSet.has(input));

      if (freshInput instanceof HTMLInputElement) {
        return freshInput;
      }

      const fallbackInput = chooseBestFileInput(inputs, context);

      if (fallbackInput instanceof HTMLInputElement) {
        return fallbackInput;
      }
    }

    return null;
  }

  function chooseBestFileInput(inputs, context) {
    if (!Array.isArray(inputs) || inputs.length === 0) {
      return null;
    }

    const rankedInputs = [...inputs].sort((left, right) => scoreFileInput(right, context) - scoreFileInput(left, context));
    return rankedInputs[0] || null;
  }

  function scoreFileInput(input, context) {
    let score = 0;

    if (context.composeRoot instanceof HTMLElement && context.composeRoot.contains(input)) {
      score += 50;
    }

    if (context.toolbar instanceof HTMLElement && context.toolbar.contains(input)) {
      score += 30;
    }

    if (input.closest('[role="dialog"]')) {
      score += 10;
    }

    if (!input.disabled) {
      score += 5;
    }

    return score;
  }

  function dispatchPointerClick(element) {
    const eventTypes = ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'];

    for (const eventType of eventTypes) {
      element.dispatchEvent(
        new MouseEvent(eventType, {
          bubbles: true,
          cancelable: true,
          view: window
        })
      );
    }
  }

  async function waitForAttachment(composeRoot, fileName) {
    const delays = [0, 160, 320, 640, 1100, 1800, 2600];

    for (const delay of delays) {
      if (delay > 0) {
        await wait(delay);
      }

      const detection = detectAttachment(composeRoot, fileName);

      if (detection.ok) {
        return detection;
      }
    }

    return {
      ok: false,
      reason: 'no-indicator'
    };
  }

  function detectAttachment(composeRoot, fileName) {
    const normalizedName = (fileName || '').toLowerCase();
    const scopes = uniqueElements([
      composeRoot,
      composeRoot ? composeRoot.parentElement : null,
      document.querySelector('main')
    ]);

    for (const scope of scopes) {
      const textContent = (scope.innerText || '').toLowerCase();

      if (normalizedName && textContent.includes(normalizedName)) {
        return {
          ok: true,
          reason: 'file-name-visible'
        };
      }

      if (scope.querySelector('[role="progressbar"]')) {
        return {
          ok: true,
          reason: 'upload-progress-visible'
        };
      }

      const removeButtons = Array.from(scope.querySelectorAll('button,[role="button"]')).some((button) => {
        const label = ((button.getAttribute('aria-label') || '') + ' ' + (button.textContent || '')).toLowerCase();
        return [
          'remover anexo',
          'remove attachment',
          'remove file',
          'remover arquivo'
        ].some((term) => label.includes(term));
      });

      if (removeButtons) {
        return {
          ok: true,
          reason: 'attachment-actions-visible'
        };
      }
    }

    return {
      ok: false,
      reason: 'no-indicator'
    };
  }

  function hasPendingAttachment(composeRoot) {
    return detectAttachment(composeRoot, '').ok;
  }

  function wait(delayMs) {
    return new Promise((resolve) => {
      window.setTimeout(resolve, delayMs);
    });
  }

  async function requestAudioStream() {
    const preferredMicrophoneId = await getPreferredMicrophoneId();
    const noiseSuppressionEnabled = await getNoiseSuppressionEnabled();

    if (preferredMicrophoneId && preferredMicrophoneId !== SYSTEM_DEFAULT_MICROPHONE_ID) {
      try {
        return await getUserMediaWithHandling({
          audio: buildPreferredMicrophoneConstraints(preferredMicrophoneId, noiseSuppressionEnabled)
        });
      } catch (error) {
        if (!isUnavailableMicrophoneError(error)) {
          throw error;
        }
      }
    }

    return getUserMediaWithHandling({
      audio: buildSystemDefaultConstraints(noiseSuppressionEnabled)
    });
  }

  function buildPreferredMicrophoneConstraints(deviceId, noiseSuppressionEnabled) {
    return {
      deviceId: { exact: deviceId },
      echoCancellation: true,
      noiseSuppression: noiseSuppressionEnabled,
      autoGainControl: true
    };
  }

  function buildSystemDefaultConstraints(noiseSuppressionEnabled) {
    return {
      deviceId: { ideal: SYSTEM_DEFAULT_MICROPHONE_ID },
      echoCancellation: true,
      noiseSuppression: noiseSuppressionEnabled,
      autoGainControl: true
    };
  }

  async function getUserMediaWithHandling(constraints) {
    try {
      return await navigator.mediaDevices.getUserMedia(constraints);
    } catch (error) {
      throw mapAudioCaptureError(error);
    }
  }

  function mapAudioCaptureError(error) {
    if (error && error.name === 'NotAllowedError') {
      return new Error(message('microphonePermissionDenied'));
    }

    if (error && (error.name === 'NotFoundError' || error.name === 'DevicesNotFoundError')) {
      return new Error(message('microphoneNotFound'));
    }

    if (error && error.name === 'OverconstrainedError') {
      return new Error(message('microphoneNotFound'));
    }

    return error instanceof Error ? error : new Error(message('startRecordingFailed'));
  }

  function isUnavailableMicrophoneError(error) {
    return error instanceof Error && error.message === message('microphoneNotFound');
  }

  async function getPreferredMicrophoneId() {
    const storage = chrome.storage?.sync;

    if (!storage) {
      return SYSTEM_DEFAULT_MICROPHONE_ID;
    }

    return new Promise((resolve) => {
      storage.get({ [STORAGE_KEY_PREFERRED_MICROPHONE_ID]: SYSTEM_DEFAULT_MICROPHONE_ID }, (items) => {
        if (chrome.runtime.lastError) {
          resolve(SYSTEM_DEFAULT_MICROPHONE_ID);
          return;
        }

        const preferredMicrophoneId = items?.[STORAGE_KEY_PREFERRED_MICROPHONE_ID];
        resolve(typeof preferredMicrophoneId === 'string' && preferredMicrophoneId ? preferredMicrophoneId : SYSTEM_DEFAULT_MICROPHONE_ID);
      });
    });
  }

  async function getNoiseSuppressionEnabled() {
    const storage = chrome.storage?.sync;

    if (!storage) {
      return DEFAULT_NOISE_SUPPRESSION_ENABLED;
    }

    return new Promise((resolve) => {
      storage.get({ [STORAGE_KEY_NOISE_SUPPRESSION_ENABLED]: DEFAULT_NOISE_SUPPRESSION_ENABLED }, (items) => {
        if (chrome.runtime.lastError) {
          resolve(DEFAULT_NOISE_SUPPRESSION_ENABLED);
          return;
        }

        resolve(items?.[STORAGE_KEY_NOISE_SUPPRESSION_ENABLED] !== false);
      });
    });
  }

  async function listAvailableMicrophones() {
    if (!navigator.mediaDevices || typeof navigator.mediaDevices.enumerateDevices !== 'function') {
      throw new Error(message('browserUnsupportedMicrophoneCapture'));
    }

    const devices = await navigator.mediaDevices.enumerateDevices();
    const audioInputs = devices
      .filter((device) => device.kind === 'audioinput')
      .filter((device) => device.deviceId && device.deviceId !== SYSTEM_DEFAULT_MICROPHONE_ID && device.deviceId !== 'communications')
      .map((device, index) => ({
        deviceId: device.deviceId,
        label: device.label || message('microphoneGenericName') + ' ' + (index + 1)
      }));

    return {
      devices: audioInputs,
      preferredMicrophoneId: await getPreferredMicrophoneId(),
      labelsAvailable: audioInputs.some((device) => !device.label.startsWith(message('microphoneGenericName')))
    };
  }

  function findActiveEditor() {
    const activeElement = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const activeCandidate = activeElement ? activeElement.closest('[contenteditable="true"], [role="textbox"], textarea') : null;

    if (isEditorCandidate(activeCandidate)) {
      return activeCandidate;
    }

    const candidates = uniqueElements(
      EDITOR_SELECTORS.flatMap((selector) => Array.from(document.querySelectorAll(selector)))
    ).filter(isEditorCandidate);

    candidates.sort((left, right) => scoreEditor(right) - scoreEditor(left));
    return candidates[0] || null;
  }

  function scoreEditor(element) {
    if (!(element instanceof HTMLElement)) {
      return 0;
    }

    const rect = element.getBoundingClientRect();
    const ariaLabel = ((element.getAttribute('aria-label') || '') + ' ' + (element.getAttribute('data-tid') || '')).toLowerCase();
    let score = 0;

    if (document.activeElement === element || element.contains(document.activeElement)) {
      score += 120;
    }

    if (element.getAttribute('role') === 'textbox') {
      score += 40;
    }

    if (/message|mensagem|compose|write|reply|chat|digite|escreva/.test(ariaLabel)) {
      score += 35;
    }

    if (rect.bottom > window.innerHeight * 0.5) {
      score += 15;
    }

    if (rect.width > 160) {
      score += 10;
    }

    if (rect.height > 24) {
      score += 10;
    }

    return score;
  }

  function findComposeRoot(editor) {
    if (!(editor instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      '[data-tid="chat-pane-compose-message-footer"]',
      '[data-tid]',
      '[role="group"]',
      'form',
      'section',
      'article',
      'main'
    ];

    for (const selector of selectors) {
      const candidate = editor.closest(selector);

      if (candidate instanceof HTMLElement && isVisible(candidate)) {
        return candidate;
      }
    }

    return editor.parentElement instanceof HTMLElement ? editor.parentElement : editor;
  }

  function isEditorCandidate(element) {
    if (!(element instanceof HTMLElement)) {
      return false;
    }

    const isTextarea = element instanceof HTMLTextAreaElement;
    const isContentEditable = element.isContentEditable || element.getAttribute('contenteditable') === 'true';
    const roleTextbox = element.getAttribute('role') === 'textbox';

    return isVisible(element) && (isTextarea || isContentEditable || roleTextbox);
  }

  function isVisible(element) {
    const rect = element.getBoundingClientRect();
    const styles = window.getComputedStyle(element);

    return rect.width > 0 && rect.height > 0 && styles.visibility !== 'hidden' && styles.display !== 'none';
  }

  function focusElement(element) {
    if (!(element instanceof HTMLElement)) {
      return;
    }

    element.focus({ preventScroll: false });
    element.click();
  }

  function chooseMimeType() {
    if (typeof MediaRecorder === 'undefined' || typeof MediaRecorder.isTypeSupported !== 'function') {
      return '';
    }

    return MIME_TYPE_CANDIDATES.find((candidate) => MediaRecorder.isTypeSupported(candidate)) || '';
  }

  function cleanupRecorder() {
    if (state.mediaRecorder) {
      state.mediaRecorder.removeEventListener('dataavailable', handleDataAvailable);
    }

    stopStream(state.mediaStream);
    clearInterval(state.timerId);
    clearTimeout(state.autoStopId);

    state.isRecording = false;
    state.mediaRecorder = null;
    state.mediaStream = null;
    state.mimeType = '';
    state.chunks = [];
    state.startedAt = 0;
    state.timerId = null;
    state.autoStopId = null;

    syncUi();
  }

  function stopStream(stream) {
    if (!stream) {
      return;
    }

    for (const track of stream.getTracks()) {
      track.stop();
    }
  }

  function startTimer() {
    clearInterval(state.timerId);
    state.timerId = window.setInterval(() => {
      if (!state.isRecording) {
        return;
      }

      showPersistentMessage(withDuration('recordingInProgressPrefix', formatDuration(Date.now() - state.startedAt)), 'recording');
    }, 1000);
  }

  function formatDuration(durationMs) {
    const totalSeconds = Math.max(0, Math.floor(durationMs / 1000));
    const minutes = String(Math.floor(totalSeconds / 60)).padStart(2, '0');
    const seconds = String(totalSeconds % 60).padStart(2, '0');
    return minutes + ':' + seconds;
  }

  function showPersistentMessage(text, tone) {
    clearTimeout(state.messageTimerId);
    state.messageText = text;
    state.messageTone = tone;
    syncUi();
  }

  function showMessage(text, tone, holdMs) {
    clearTimeout(state.messageTimerId);
    state.messageText = text;
    state.messageTone = tone;
    syncUi();

    state.messageTimerId = window.setTimeout(() => {
      if (!state.isRecording && !state.isBusy) {
        clearMessage();
      }
    }, holdMs);
  }

  function clearMessage() {
    clearTimeout(state.messageTimerId);
    state.messageText = '';
    state.messageTone = 'info';
    syncUi();
  }

  function normalizeError(error) {
    if (!(error instanceof Error)) {
      return message('unexpectedVoiceMessageError');
    }

    return error.message;
  }

  function syncUi() {
    if (state.isContextInvalidated) {
      return;
    }

    const root = document.getElementById(ROOT_ID);
    const button = document.getElementById(BUTTON_ID);
    const bubble = ensureBubble();

    if (root instanceof HTMLElement && button instanceof HTMLButtonElement) {
      const stateName = state.isRecording ? 'recording' : state.isBusy ? 'busy' : 'idle';
      root.dataset.state = stateName;
      button.disabled = state.isBusy;
      button.setAttribute('aria-pressed', String(state.isRecording));
      button.setAttribute(
        'aria-label',
        state.isRecording ? message('stopButtonLabel') : message('recordButtonLabel')
      );
      button.setAttribute(
        'title',
        state.isRecording ? message('stopButtonLabel') : message('recordButtonLabel')
      );
    }

    if (!(bubble instanceof HTMLElement)) {
      return;
    }

    if (state.messageText && button instanceof HTMLButtonElement) {
      bubble.hidden = false;
      bubble.textContent = state.messageText;
      bubble.dataset.tone = state.messageTone;
      positionBubble(button, bubble);
    } else {
      bubble.hidden = true;
      bubble.textContent = '';
      bubble.style.removeProperty('top');
      bubble.style.removeProperty('left');
      delete bubble.dataset.tone;
    }
  }

  function positionBubble(button, bubble) {
    const rect = button.getBoundingClientRect();

    if (rect.width <= 0 || rect.height <= 0) {
      bubble.hidden = true;
      return;
    }

    const bubbleRect = bubble.getBoundingClientRect();
    const margin = 8;
    let top = rect.top - bubbleRect.height - 10;

    if (top < margin) {
      top = rect.bottom + 10;
    }

    const maxLeft = Math.max(margin, window.innerWidth - bubbleRect.width - margin);
    const left = Math.max(margin, Math.min(rect.right - bubbleRect.width, maxLeft));

    bubble.style.top = Math.round(top) + 'px';
    bubble.style.left = Math.round(left) + 'px';
  }

  function uniqueElements(elements) {
    const unique = [];
    const seen = new Set();

    for (const element of elements) {
      if (!(element instanceof HTMLElement)) {
        continue;
      }

      if (seen.has(element)) {
        continue;
      }

      seen.add(element);
      unique.push(element);
    }

    return unique;
  }

  function runtimeUrl(path) {
    if (state.isContextInvalidated) {
      return path;
    }

    try {
      return chrome.runtime.getURL(path);
    } catch (error) {
      if (isExtensionContextInvalidated(error)) {
        invalidateExtensionContext();
        return path;
      }

      throw error;
    }
  }

  function isExtensionContextInvalidated(error) {
    return error instanceof Error && /extension context invalidated/i.test(error.message);
  }

  function invalidateExtensionContext() {
    if (state.isContextInvalidated) {
      return;
    }

    state.isContextInvalidated = true;

    if (state.observer) {
      state.observer.disconnect();
      state.observer = null;
    }

    if (state.runtimeMessageListener && chrome.runtime?.onMessage) {
      try {
        chrome.runtime.onMessage.removeListener(state.runtimeMessageListener);
      } catch (_error) {
      }
    }

    if (state.refreshRaf) {
      window.cancelAnimationFrame(state.refreshRaf);
      state.refreshRaf = 0;
    }

    clearInterval(state.timerId);
    clearTimeout(state.autoStopId);
    clearTimeout(state.messageTimerId);
    stopStream(state.mediaStream);

    state.isBusy = false;
    state.isRecording = false;
    state.mediaRecorder = null;
    state.mediaStream = null;
    state.messageText = '';
    state.pendingAttachmentName = '';
    state.runtimeMessageListener = null;

    document.getElementById(ROOT_ID)?.remove();
    document.getElementById(BUBBLE_ID)?.remove();
    window.removeEventListener('resize', scheduleRefresh, { passive: true });
    window.removeEventListener('focusin', scheduleRefresh, true);
    window.__teamsVoiceMessageInitialized = false;
    window.__teamsVoiceMessageRuntimeListenersInstalled = false;

    if (window[INSTANCE_KEY]?.dispose === invalidateExtensionContext) {
      delete window[INSTANCE_KEY];
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrap, { once: true });
  } else {
    bootstrap();
  }
})();