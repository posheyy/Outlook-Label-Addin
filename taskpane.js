/* Mail-Labels – Outlook-Taskpane
 * Oberfläche für die Label-Konfiguration. Die gesamte Automatisierung läuft in n8n:
 *   - Workflow "Mail-Labels API (Outlook-Add-in)"   → Labels lesen/anlegen/ändern/löschen, Testklassifizierung
 *   - Workflow "Mail-Klassifizierung PoC"           → klassifiziert Mails und setzt Outlook-Kategorien
 */
(function () {
  'use strict';

  var DEFAULT_URL = 'https://n8n.cpronect.de/webhook/mail-labels';
  var API_KEY_HEADER = 'X-Api-Key';
  var SETTINGS_URL = 'mailLabels.url';
  var SETTINGS_KEY = 'mailLabels.apiKey';

  // Outlook-Farbpresets (Office.MailboxEnums.CategoryColor)
  var COLORS = [
    { id: 'Preset0', name: 'Rot', hex: '#e74856' },
    { id: 'Preset1', name: 'Orange', hex: '#ff8c00' },
    { id: 'Preset2', name: 'Braun', hex: '#ab8b5e' },
    { id: 'Preset3', name: 'Gelb', hex: '#fff100' },
    { id: 'Preset4', name: 'Grün', hex: '#47d041' },
    { id: 'Preset5', name: 'Türkis', hex: '#30c6cc' },
    { id: 'Preset6', name: 'Oliv', hex: '#73aa24' },
    { id: 'Preset7', name: 'Blau', hex: '#4cb3ff' },
    { id: 'Preset8', name: 'Lila', hex: '#a066e0' },
    { id: 'Preset9', name: 'Cranberry', hex: '#ee5fb7' },
    { id: 'Preset10', name: 'Stahlblau', hex: '#98b6c8' },
    { id: 'Preset11', name: 'Dunkles Stahlblau', hex: '#4f6b78' },
    { id: 'Preset12', name: 'Grau', hex: '#ababab' },
    { id: 'Preset13', name: 'Dunkelgrau', hex: '#666666' },
    { id: 'Preset14', name: 'Schwarz', hex: '#474747' },
    { id: 'Preset15', name: 'Dunkelrot', hex: '#910a19' },
    { id: 'Preset16', name: 'Dunkelorange', hex: '#ce4b28' },
    { id: 'Preset17', name: 'Dunkelbraun', hex: '#99694c' },
    { id: 'Preset18', name: 'Dunkelgelb', hex: '#b0a923' },
    { id: 'Preset19', name: 'Dunkelgrün', hex: '#027400' },
    { id: 'Preset20', name: 'Dunkeltürkis', hex: '#1e8287' },
    { id: 'Preset21', name: 'Dunkeloliv', hex: '#46651e' },
    { id: 'Preset22', name: 'Dunkelblau', hex: '#2f5596' },
    { id: 'Preset23', name: 'Dunkellila', hex: '#5f3c94' },
    { id: 'Preset24', name: 'Dunkelcranberry', hex: '#9c2c64' }
  ];

  // Systemlabels: werden vom n8n-Workflow vergeben, im Add-in nicht editierbar
  var SYSTEM_CATEGORIES = [
    { name: 'Sonstiges', farbe: 'Preset13' },
    { name: 'Prüfen', farbe: 'Preset3' },
    { name: 'AI-klassifiziert', farbe: 'Preset12' }
  ];
  var RESERVED = SYSTEM_CATEGORIES.map(function (c) { return c.name.toLowerCase(); });

  var state = { labels: [], editing: null, mailbox: '', confirmDeleteId: null };

  function $(id) { return document.getElementById(id); }

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (k === 'text') node.textContent = attrs[k];
      else if (k === 'class') node.className = attrs[k];
      else if (k.indexOf('on') === 0) node.addEventListener(k.slice(2), attrs[k]);
      else node.setAttribute(k, attrs[k]);
    });
    (children || []).forEach(function (c) { if (c) node.appendChild(c); });
    return node;
  }

  function colorHex(id) {
    var c = COLORS.filter(function (x) { return x.id === id; })[0];
    return c ? c.hex : '#cccccc';
  }

  // ---------- Status ----------
  var statusTimer = null;
  function showStatus(msg, type, sticky) {
    var s = $('status');
    s.textContent = msg;
    s.className = 'status ' + (type || 'info');
    clearTimeout(statusTimer);
    if (!sticky) statusTimer = setTimeout(function () { s.className = 'status hidden'; }, 4000);
  }
  function hideStatus() { $('status').className = 'status hidden'; }

  // ---------- Einstellungen (pro Postfach, roamt mit) ----------
  function getSetting(k) { return Office.context.roamingSettings.get(k); }
  function saveSettings(url, key) {
    return new Promise(function (resolve, reject) {
      Office.context.roamingSettings.set(SETTINGS_URL, url);
      Office.context.roamingSettings.set(SETTINGS_KEY, key);
      Office.context.roamingSettings.saveAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(r.error && r.error.message));
      });
    });
  }
  function apiUrl() { return getSetting(SETTINGS_URL) || DEFAULT_URL; }
  function apiKey() { return getSetting(SETTINGS_KEY) || ''; }

  // ---------- API zu n8n ----------
  function api(action, payload) {
    if (!apiKey()) {
      showView('settings');
      return Promise.reject(new Error('Bitte zuerst den API-Key hinterlegen.'));
    }
    var headers = { 'Content-Type': 'application/json' };
    headers[API_KEY_HEADER] = apiKey();
    var body = Object.assign({ action: action, mailbox: state.mailbox }, payload || {});
    return fetch(apiUrl(), { method: 'POST', headers: headers, body: JSON.stringify(body) })
      .then(function (res) {
        return res.text().then(function (txt) {
          var data = null;
          try { data = txt ? JSON.parse(txt) : null; } catch (e) { data = null; }
          if (res.status === 401 || res.status === 403) throw new Error('API-Key abgelehnt (' + res.status + ').');
          if (res.status === 404) throw new Error('Webhook nicht gefunden – ist der n8n-Workflow aktiv?');
          if (!res.ok || !data || data.ok === false) {
            throw new Error((data && data.error) || ('Fehler ' + res.status + ' von n8n.'));
          }
          return data;
        });
      }, function () {
        throw new Error('n8n nicht erreichbar (Netzwerk oder CORS).');
      });
  }

  // ---------- Outlook-Kategorien (Master-Liste) ----------
  function getMasterCategories() {
    return new Promise(function (resolve) {
      var mc = Office.context.mailbox.masterCategories;
      if (!mc) return resolve(null);
      mc.getAsync(function (r) { resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null); });
    });
  }
  function ensureCategories(items) {
    return getMasterCategories().then(function (existing) {
      if (!existing) return false;
      var names = existing.map(function (c) { return c.displayName.toLowerCase(); });
      var missing = items
        .filter(function (i) { return names.indexOf(i.name.toLowerCase()) === -1; })
        .map(function (i) { return { displayName: i.name, color: i.farbe }; });
      if (!missing.length) return true;
      return new Promise(function (resolve) {
        Office.context.mailbox.masterCategories.addAsync(missing, function (r) {
          resolve(r.status === Office.AsyncResultStatus.Succeeded);
        });
      });
    });
  }

  // ---------- Views ----------
  function showView(name) {
    $('viewList').classList.toggle('hidden', name !== 'list');
    $('viewForm').classList.toggle('hidden', name !== 'form');
    $('viewSettings').classList.toggle('hidden', name !== 'settings');
    if (name === 'settings') {
      $('setUrl').value = apiUrl();
      $('setKey').value = apiKey();
    }
  }

  // ---------- Liste ----------
  // keepStatus: true nach Speichern/Löschen, damit die Erfolgsmeldung stehen bleibt
  function loadLabels(keepStatus) {
    if (!keepStatus) showStatus('Lade Labels …', 'info', true);
    return api('list').then(function (data) {
      state.labels = data.labels || [];
      renderList();
      if (!keepStatus) hideStatus();
      ensureCategories(state.labels.concat(SYSTEM_CATEGORIES));
    }).catch(function (e) { showStatus(e.message, 'error', true); });
  }

  function descriptionWarning(l) {
    var len = (l.beschreibung || '').trim().length;
    if (len === 0) return 'Keine Beschreibung – Mistral entscheidet nur nach dem Namen.';
    if (len < 40) return 'Sehr kurze Beschreibung – Abgrenzung ergänzen.';
    return '';
  }

  function renderList() {
    var ul = $('labelList');
    ul.innerHTML = '';
    if (!state.labels.length) {
      ul.appendChild(el('li', { class: 'muted', text: 'Noch keine Labels. Lege dein erstes Label an.' }));
      return;
    }
    state.labels.forEach(function (l) {
      var warn = descriptionWarning(l);
      var li = el('li', { class: l.aktiv ? '' : 'inactive' }, [
        el('div', { class: 'head' }, [
          el('span', { class: 'dot', style: 'background:' + colorHex(l.farbe) }),
          el('span', { class: 'name', text: l.name }),
          l.aktiv ? null : el('span', { class: 'badge', text: 'inaktiv' })
        ]),
        l.beschreibung ? el('div', { class: 'desc', text: l.beschreibung }) : null,
        warn ? el('div', { class: 'warn', text: '⚠ ' + warn }) : null
      ]);

      if (state.confirmDeleteId === l.id) {
        li.appendChild(el('div', { class: 'confirm' }, [
          el('div', { text: 'Label „' + l.name + '“ löschen? Bereits getaggte Mails behalten die Kategorie, neue Mails bekommen dieses Label nicht mehr.' }),
          el('div', { class: 'actions' }, [
            el('button', { class: 'danger', text: 'Ja, löschen', onclick: function () { deleteLabel(l); } }),
            el('button', { text: 'Abbrechen', onclick: function () { state.confirmDeleteId = null; renderList(); } })
          ])
        ]));
      } else {
        li.appendChild(el('div', { class: 'actions' }, [
          el('button', { text: 'Bearbeiten', onclick: function () { openForm(l); } }),
          el('button', { text: l.aktiv ? 'Deaktivieren' : 'Aktivieren', onclick: function () { toggleActive(l); } }),
          el('button', { class: 'danger', text: 'Löschen', onclick: function () { state.confirmDeleteId = l.id; renderList(); } })
        ]));
      }
      ul.appendChild(li);
    });
  }

  function toggleActive(l) {
    var payload = Object.assign({}, l, { aktiv: !l.aktiv });
    api('update', payload).then(function () {
      showStatus('„' + l.name + '“ ' + (payload.aktiv ? 'aktiviert' : 'deaktiviert') + '.', 'ok');
      return loadLabels(true);
    }).catch(function (e) { showStatus(e.message, 'error', true); });
  }

  function deleteLabel(l) {
    state.confirmDeleteId = null;
    api('delete', { id: l.id }).then(function () {
      showStatus('„' + l.name + '“ gelöscht.', 'ok');
      return loadLabels(true);
    }).catch(function (e) { showStatus(e.message, 'error', true); renderList(); });
  }

  // ---------- Formular ----------
  function fillColors() {
    var sel = $('fFarbe');
    COLORS.forEach(function (c) { sel.appendChild(el('option', { value: c.id, text: c.name })); });
    sel.addEventListener('change', updateColorPreview);
  }
  function updateColorPreview() { $('colorPreview').style.background = colorHex($('fFarbe').value); }

  function updateQuality() {
    var d = $('fBeschreibung').value.trim();
    var b = $('fBeispiele').value.trim();
    var q = $('qualityHint');
    var hasNegation = /\b(nicht|kein|keine|ausgenommen|außer)\b/i.test(d);
    if (d.length < 40) {
      q.className = 'quality weak';
      q.textContent = 'Zu knapp: Beschreibe, woran man eine solche Mail erkennt.';
    } else if (!hasNegation || !b) {
      q.className = 'quality medium';
      q.textContent = 'Ok. Besser wird es mit ' +
        (!hasNegation ? 'einer Abgrenzung („NICHT darunter fällt …“)' : '') +
        (!hasNegation && !b ? ' und ' : '') +
        (!b ? 'typischen Begriffen/Beispielen' : '') + '.';
    } else {
      q.className = 'quality good';
      q.textContent = 'Gute Beschreibung: Erkennungsmerkmale, Abgrenzung und Beispiele vorhanden.';
    }
  }

  function openForm(label) {
    state.editing = label || null;
    hideStatus();
    $('formTitle').textContent = label ? 'Label bearbeiten' : 'Neues Label';
    $('fName').value = label ? label.name : '';
    $('fBeschreibung').value = label ? label.beschreibung : '';
    $('fBeispiele').value = label ? label.beispiele : '';
    $('fFarbe').value = label ? label.farbe : 'Preset7';
    $('fAktiv').checked = label ? label.aktiv : true;
    $('renameNote').classList.add('hidden');
    updateColorPreview();
    updateQuality();
    showView('form');
    $('fName').focus();
  }

  function saveForm() {
    var name = $('fName').value.trim();
    var data = {
      name: name,
      beschreibung: $('fBeschreibung').value.trim(),
      beispiele: $('fBeispiele').value.trim(),
      farbe: $('fFarbe').value,
      aktiv: $('fAktiv').checked
    };
    if (!name) return showStatus('Name darf nicht leer sein.', 'error');
    if (RESERVED.indexOf(name.toLowerCase()) !== -1) return showStatus('„' + name + '“ ist ein Systemlabel.', 'error');
    var dup = state.labels.some(function (l) {
      return l.name.toLowerCase() === name.toLowerCase() && (!state.editing || l.id !== state.editing.id);
    });
    if (dup) return showStatus('Ein Label mit diesem Namen gibt es schon.', 'error');

    var isEdit = !!state.editing;
    var renamed = isEdit && state.editing.name !== name;
    if (isEdit) data.id = state.editing.id;

    $('btnSave').disabled = true;
    api(isEdit ? 'update' : 'create', data).then(function () {
      return ensureCategories([{ name: name, farbe: data.farbe }]);
    }).then(function () {
      var msg = isEdit ? 'Label gespeichert.' : 'Label angelegt.';
      if (renamed) msg += ' Alte Mails behalten die Kategorie „' + state.editing.name + '“.';
      showStatus(msg, 'ok');
      state.editing = null;
      showView('list');
      return loadLabels(true);
    }).catch(function (e) {
      showStatus(e.message, 'error', true);
    }).then(function () { $('btnSave').disabled = false; });
  }

  // ---------- Test an der geöffneten Mail ----------
  function readBodyText(item) {
    return new Promise(function (resolve) {
      item.body.getAsync(Office.CoercionType.Text, function (r) {
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
      });
    });
  }

  function runTest() {
    var item = Office.context.mailbox.item;
    var box = $('testResult');
    if (!item) { showStatus('Keine Mail geöffnet.', 'error'); return; }
    $('btnTest').disabled = true;
    box.className = '';
    box.textContent = 'Klassifiziere …';
    readBodyText(item).then(function (text) {
      return api('test', {
        subject: item.subject || '',
        sender: (item.from && item.from.emailAddress) || '',
        body: (text || '').slice(0, 6000)
      });
    }).then(function (data) {
      var r = data.result || {};
      var label = state.labels.filter(function (l) { return l.name === r.tag; })[0];
      var conf = Math.round((r.confidence || 0) * 100);
      box.innerHTML = '';
      box.appendChild(el('div', { class: 'tag' }, [
        el('span', { class: 'dot', style: 'background:' + colorHex(label ? label.farbe : 'Preset13') }),
        el('span', { text: r.tag || '–' }),
        el('span', { class: 'badge', text: conf + ' %' })
      ]));
      if (conf < 60) box.appendChild(el('div', { class: 'warn small', text: 'Unsicher – würde zusätzlich „Prüfen“ bekommen.' }));
      box.appendChild(el('div', { class: 'muted small', text: r.reason || '' }));
    }).catch(function (e) {
      box.textContent = '';
      showStatus(e.message, 'error', true);
    }).then(function () { $('btnTest').disabled = false; });
  }

  // ---------- Init ----------
  function init() {
    state.mailbox = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();
    $('mailbox').textContent = state.mailbox;
    fillColors();

    $('btnNew').addEventListener('click', function () { openForm(null); });
    $('btnCancel').addEventListener('click', function () { state.editing = null; showView('list'); });
    $('btnSave').addEventListener('click', saveForm);
    $('btnTest').addEventListener('click', runTest);
    $('fBeschreibung').addEventListener('input', updateQuality);
    $('fBeispiele').addEventListener('input', updateQuality);
    $('fName').addEventListener('input', function () {
      var renamed = state.editing && $('fName').value.trim() !== state.editing.name;
      $('renameNote').classList.toggle('hidden', !renamed);
    });

    $('btnSettings').addEventListener('click', function () { showView('settings'); });
    $('btnCancelSettings').addEventListener('click', function () { showView('list'); });
    $('btnSaveSettings').addEventListener('click', function () {
      var url = $('setUrl').value.trim() || DEFAULT_URL;
      var key = $('setKey').value.trim();
      if (!key) return showStatus('API-Key fehlt.', 'error');
      saveSettings(url, key).then(function () {
        showStatus('Verbindung gespeichert.', 'ok');
        showView('list');
        loadLabels();
      }).catch(function (e) { showStatus('Speichern fehlgeschlagen: ' + e.message, 'error', true); });
    });

    // Angeheftete Taskpane: Testergebnis zurücksetzen, wenn eine andere Mail geöffnet wird
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        $('testResult').className = 'hidden';
      });
    }

    if (!apiKey()) {
      showView('settings');
      showStatus('Einmalig: n8n-Webhook-URL und API-Key eintragen.', 'info', true);
    } else {
      showView('list');
      loadLabels();
    }
  }

  Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) init();
  });
})();
