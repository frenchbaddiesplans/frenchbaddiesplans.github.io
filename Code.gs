/**
 * BaddiesPlans — prise de rendez-vous pour French Baddies Nails
 * Phases 2 + 3 : socle (Google Sheets + Agenda) et page publique de réservation.
 *
 * Tout se règle dans le Google Sheets qui contient ce script :
 *   - Réglages   : textes, adresse, pas entre créneaux, pause, délai minimum, etc.
 *   - Prestations: nom, durée, prix, description, active OUI/NON
 *   - Horaires   : une ligne par plage d'ouverture (plusieurs lignes possibles par jour)
 *   - Questions  : infos demandées au visiteur
 *   - Réservations : historique (rempli automatiquement)
 * Pour bloquer un créneau ou une journée : créer un événement dans Google Agenda.
 */

var TZ = 'Europe/Paris';
var TAG = 'baddiesplans';
var JOURS = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
var MOIS = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];

/* ------------------------------------------------------------------ */
/* Installation : à lancer une fois depuis l'éditeur (bouton Exécuter) */
/* ------------------------------------------------------------------ */

function installer() {
  var ss = SpreadsheetApp.getActive();
  creerOnglet_(ss, 'Réglages', ['Clé', 'Valeur', 'À quoi ça sert'], [
    ['nom_salon', 'French Baddies Nails', 'Nom affiché en haut de la page'],
    ['slogan', 'Tu veux des ongles de folie ?', 'Phrase d\'accroche'],
    ['presentation', 'Passionnée d\'onglerie et diplômée d\'esthétique, je t\'accueille à Caen pour tes manucures et poses américaines.', 'Texte de présentation'],
    ['lieu', 'Fripe à la mode de Caen', 'Nom du lieu affiché'],
    ['adresse', '17 rue de la Pigacière, 14000 Caen', 'Adresse affichée et envoyée dans la confirmation'],
    ['pas_minutes', '30', 'Écart entre deux heures de début proposées (minutes)'],
    ['pause_minutes', '15', 'Pause minimum entre deux rendez-vous (minutes)'],
    ['delai_min_heures', '24', 'On ne peut pas réserver moins de X heures à l\'avance'],
    ['jours_visibles', '21', 'Nombre de jours affichés à l\'avance'],
    ['max_rdv_par_jour', '3', 'Nombre maximum de rendez-vous par jour'],
    ['a_savoir', 'J\'ai un chien avec moi, il est très gentil et dort la plupart du temps dans le canapé.\nPour une pose avec nail art, il faut prévoir un créneau de 3 h : je préfère prendre mon temps que bâcler.\nJe ne pratique pas sur les ongles très rongés ni sur les personnes de moins de 16 ans.', 'Bloc « À savoir » (une ligne = un point)'],
    ['message_confirmation', 'Merci pour ta réservation ! Si tu as un empêchement, préviens-moi au moins 24 h avant.', 'Message ajouté à l\'e-mail de confirmation'],
    ['me_notifier', 'OUI', 'OUI = recevoir un e-mail à chaque réservation']
  ]);
  creerOnglet_(ss, 'Prestations', ['Nom', 'Durée (min)', 'Prix', 'Description', 'Active (OUI/NON)'], [
    ['Pose américaine', '120', '', 'Pose complète', 'OUI'],
    ['Pose américaine + nail art', '180', '', 'Prévoir 3 h pour un résultat parfait', 'OUI'],
    ['Manucure', '60', '', 'Soin des mains et des ongles', 'OUI']
  ]);
  creerOnglet_(ss, 'Horaires', ['Jour', 'Début (HH:MM)', 'Fin (HH:MM)'], [
    ['mardi', '10:00', '18:00'],
    ['mercredi', '10:00', '18:00'],
    ['jeudi', '10:00', '18:00'],
    ['vendredi', '10:00', '18:00'],
    ['samedi', '10:00', '17:00']
  ]);
  creerOnglet_(ss, 'Questions', ['Question', 'Type (texte, email, tel, long, case)', 'Obligatoire (OUI/NON)'], [
    ['Prénom et nom', 'texte', 'OUI'],
    ['E-mail', 'email', 'OUI'],
    ['Téléphone', 'tel', 'OUI'],
    ['Ton idée de design (couleurs, inspi, lien Pinterest…)', 'long', 'NON'],
    ['J\'ai 16 ans ou plus', 'case', 'OUI'],
    ['Mes ongles ne sont pas très rongés', 'case', 'OUI']
  ]);
  creerOnglet_(ss, 'Réservations', ['Créée le', 'Début', 'Fin', 'Prestation', 'Nom', 'E-mail', 'Téléphone', 'Réponses', 'ID événement', 'Statut'], []);
  var defaut = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
  if (defaut && ss.getSheets().length > 1) ss.deleteSheet(defaut);
  if (Session.getScriptTimeZone() !== TZ) {
    throw new Error('Fuseau horaire du script : ' + Session.getScriptTimeZone() + '. Passe-le sur Europe/Paris (Paramètres du projet), puis relance installer().');
  }
  CalendarApp.getDefaultCalendar(); // déclenche l'autorisation Agenda
  return 'Installation terminée';
}

function creerOnglet_(ss, nom, entetes, lignes) {
  if (ss.getSheetByName(nom)) return; // ne jamais écraser ce qui existe
  var sh = ss.insertSheet(nom);
  var data = [entetes].concat(lignes);
  var range = sh.getRange(1, 1, data.length, entetes.length);
  range.setNumberFormat('@'); // texte brut : évite les conversions d'heures/dates
  range.setValues(data);
  sh.getRange(1, 1, 1, entetes.length).setFontWeight('bold').setBackground('#ff5a3c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, Math.max(data.length, 50), entetes.length).setNumberFormat('@').setWrap(true);
  sh.autoResizeColumns(1, entetes.length);
}

/* ---------------------------- Lecture des réglages ---------------------------- */

function lignes_(nom) {
  var sh = SpreadsheetApp.getActive().getSheetByName(nom);
  if (!sh) throw new Error('Onglet manquant : ' + nom + ' (lance installer())');
  var v = sh.getDataRange().getDisplayValues();
  v.shift();
  return v.filter(function (r) { return r.join('').trim() !== ''; });
}

function reglages_() {
  var o = {};
  lignes_('Réglages').forEach(function (r) { o[String(r[0]).trim()] = String(r[1]); });
  return {
    nom: o.nom_salon || 'French Baddies Nails',
    slogan: o.slogan || '',
    presentation: o.presentation || '',
    lieu: o.lieu || '',
    adresse: o.adresse || '',
    pas: Math.max(5, parseInt(o.pas_minutes, 10) || 30),
    pause: Math.max(0, parseInt(o.pause_minutes, 10) || 0),
    delaiH: Math.max(0, parseFloat(String(o.delai_min_heures).replace(',', '.')) || 0),
    jours: Math.min(90, Math.max(1, parseInt(o.jours_visibles, 10) || 21)),
    maxJour: parseInt(o.max_rdv_par_jour, 10) || 0,
    aSavoir: (o.a_savoir || '').split('\n').map(function (s) { return s.trim(); }).filter(String),
    messageConfirmation: o.message_confirmation || '',
    meNotifier: /^oui$/i.test(String(o.me_notifier).trim())
  };
}

function prestations_() {
  return lignes_('Prestations')
    .map(function (r, i) {
      return { id: i, nom: r[0].trim(), duree: parseInt(r[1], 10) || 0, prix: r[2].trim(), description: r[3].trim(), active: !/^non$/i.test(r[4].trim()) };
    })
    .filter(function (p) { return p.nom && p.duree > 0 && p.active; });
}

function questions_() {
  return lignes_('Questions').map(function (r, i) {
    var type = (r[1] || 'texte').trim().toLowerCase();
    if (['texte', 'email', 'tel', 'long', 'case'].indexOf(type) < 0) type = 'texte';
    return { id: 'q' + i, label: r[0].trim(), type: type, obligatoire: /^oui$/i.test(r[2].trim()) };
  }).filter(function (q) { return q.label; });
}

function horaires_() {
  var parJour = {};
  lignes_('Horaires').forEach(function (r) {
    var j = JOURS.indexOf(r[0].trim().toLowerCase());
    var d = minutes_(r[1]), f = minutes_(r[2]);
    if (j < 0 || d === null || f === null || f <= d) return;
    (parJour[j] = parJour[j] || []).push([d, f]);
  });
  return parJour;
}

function minutes_(s) {
  var m = String(s).trim().match(/^(\d{1,2})\s*[:hH]\s*(\d{2})?/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + (m[2] ? parseInt(m[2], 10) : 0);
}

/* ------------------------------- Page web ------------------------------- */

function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.logo = LOGO_DATA_URI;
  t.boot = JSON.stringify(donneesPage()).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('French Baddies Nails — Prendre rendez-vous')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function donneesPage() {
  var r = reglages_();
  return {
    salon: { nom: r.nom, slogan: r.slogan, presentation: r.presentation, lieu: r.lieu, adresse: r.adresse, aSavoir: r.aSavoir },
    prestations: prestations_(),
    questions: questions_()
  };
}

/* ----------------------------- Créneaux ----------------------------- */

/** Créneaux d'une prestation sur la période visible : libres et pris. */
function getCreneaux(prestationId) {
  var r = reglages_();
  var p = prestations_().filter(function (x) { return x.id === Number(prestationId); })[0];
  if (!p) throw new Error('Prestation introuvable');
  return calculerCreneaux_(r, p, horaires_(), new Date());
}

function calculerCreneaux_(r, p, horaires, maintenant) {
  var debutPeriode = new Date(maintenant.getFullYear(), maintenant.getMonth(), maintenant.getDate());
  var finPeriode = new Date(debutPeriode.getTime() + r.jours * 864e5 + 864e5);
  var events = CalendarApp.getDefaultCalendar().getEvents(debutPeriode, finPeriode);
  var occupe = [], rdvParJour = {};
  events.forEach(function (ev) {
    if (ev.isAllDayEvent()) {
      occupe.push([ev.getAllDayStartDate().getTime(), ev.getAllDayEndDate().getTime()]);
      return;
    }
    var s = ev.getStartTime().getTime(), e = ev.getEndTime().getTime();
    occupe.push([s - r.pause * 6e4, e + r.pause * 6e4]);
    if (ev.getTag(TAG) && ev.getTag(TAG + '_statut') !== 'annule') {
      var k = cleJour_(ev.getStartTime());
      rdvParJour[k] = (rdvParJour[k] || 0) + 1;
    }
  });
  var limite = maintenant.getTime() + r.delaiH * 36e5;
  var jours = [];
  for (var i = 0; i < r.jours; i++) {
    var jour = new Date(debutPeriode.getFullYear(), debutPeriode.getMonth(), debutPeriode.getDate() + i);
    var plages = horaires[jour.getDay()] || [];
    var complet = r.maxJour > 0 && (rdvParJour[cleJour_(jour)] || 0) >= r.maxJour;
    var slots = [];
    plages.forEach(function (pl) {
      for (var m = pl[0]; m + p.duree <= pl[1]; m += r.pas) {
        var s = new Date(jour.getFullYear(), jour.getMonth(), jour.getDate(), Math.floor(m / 60), m % 60).getTime();
        if (s < limite) continue;
        var e = s + p.duree * 6e4;
        var pris = complet || occupe.some(function (o) { return o[0] < e && o[1] > s; });
        slots.push({ heure: pad_(Math.floor(m / 60)) + ':' + pad_(m % 60), debut: s, libre: !pris });
      }
    });
    if (!plages.length) continue;
    slots.sort(function (a, b) { return a.debut - b.debut; });
    jours.push({ cle: cleJour_(jour), jourSemaine: JOURS[jour.getDay()], jour: jour.getDate(), mois: MOIS[jour.getMonth()], slots: slots });
  }
  return jours;
}

function cleJour_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }

/* ----------------------------- Réservation ----------------------------- */

function reserver(demande) {
  var r = reglages_();
  var p = prestations_().filter(function (x) { return x.id === Number(demande.prestationId); })[0];
  if (!p) throw new Error('Cette prestation n\'est plus disponible.');
  var qs = questions_();
  var rep = demande.reponses || {};
  var erreurs = [];
  qs.forEach(function (q) {
    var v = rep[q.id];
    if (q.type === 'case') {
      if (q.obligatoire && v !== true) erreurs.push('Merci de cocher : « ' + q.label + ' »');
      return;
    }
    v = String(v || '').trim().slice(0, 2000);
    rep[q.id] = v;
    if (q.obligatoire && !v) erreurs.push('Champ manquant : ' + q.label);
    if (v && q.type === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) erreurs.push('E-mail invalide');
  });
  if (erreurs.length) throw new Error(erreurs.join('\n'));

  var champ = function (type) { var q = qs.filter(function (x) { return x.type === type; })[0]; return q ? rep[q.id] : ''; };
  var nom = champ('texte') || 'Client·e', email = champ('email'), tel = champ('tel');
  var debut = new Date(Number(demande.debut));
  var fin = new Date(debut.getTime() + p.duree * 6e4);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Beaucoup de monde en même temps, réessaie dans un instant.');
  try {
    // Re-vérifie le créneau au moment même de la réservation (anti-double réservation)
    var jours = calculerCreneaux_(r, p, horaires_(), new Date());
    var ok = jours.some(function (j) { return j.slots.some(function (s) { return s.debut === debut.getTime() && s.libre; }); });
    if (!ok) throw new Error('Oups, ce créneau vient d\'être pris. Choisis-en un autre.');

    var details = qs.map(function (q) {
      var v = rep[q.id];
      return q.label + ' : ' + (q.type === 'case' ? (v === true ? 'oui' : 'non') : (v || '—'));
    }).join('\n');
    var ev = CalendarApp.getDefaultCalendar().createEvent('💅 ' + p.nom + ' — ' + nom, debut, fin, {
      description: details + '\n\nRéservé en ligne via BaddiesPlans',
      location: r.adresse
    });
    ev.setTag(TAG, '1');
    ev.removeAllReminders();
    ev.addPopupReminder(60);
    SpreadsheetApp.getActive().getSheetByName('Réservations').appendRow([
      Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'),
      Utilities.formatDate(debut, TZ, 'dd/MM/yyyy HH:mm'),
      Utilities.formatDate(fin, TZ, 'dd/MM/yyyy HH:mm'),
      p.nom, nom, email, tel, details, ev.getId(), 'confirmé'
    ]);
  } finally {
    lock.releaseLock();
  }

  var quand = dateLongue_(debut);
  try { envoyerConfirmation_(r, p, nom, email, quand); } catch (e) { console.error(e); }
  if (r.meNotifier) {
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: 'Nouveau RDV : ' + p.nom + ' — ' + quand,
        body: 'Nouvelle réservation via BaddiesPlans\n\n' + p.nom + '\n' + quand + '\n\n' + nom + '\n' + email + '\n' + tel,
        name: 'BaddiesPlans'
      });
    } catch (e) { console.error(e); }
  }
  return { prestation: p.nom, quand: quand, adresse: r.adresse, lieu: r.lieu, email: email };
}

function dateLongue_(d) {
  var j = JOURS[Number(Utilities.formatDate(d, TZ, 'u')) % 7];
  return j.charAt(0).toUpperCase() + j.slice(1) + ' ' + Utilities.formatDate(d, TZ, 'd') + ' ' +
    MOIS[Number(Utilities.formatDate(d, TZ, 'M')) - 1] + ' à ' + Utilities.formatDate(d, TZ, "HH'h'mm");
}

function envoyerConfirmation_(r, p, nom, email, quand) {
  if (!email) return;
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;border-radius:18px;overflow:hidden;border:1px solid #ffd0e6">' +
    '<div style="background:linear-gradient(135deg,#ff6a3d,#ff0a5e);color:#fff;padding:22px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold">' + esc(r.nom) + '</div><div style="opacity:.9">Ton rendez-vous est confirmé ✦</div></div>' +
    '<div style="padding:22px;color:#3a0a24;line-height:1.5">Hello ' + esc(nom) + ',<br><br>' +
    '<b>' + esc(p.nom) + '</b><br>' + esc(quand) + ' (' + p.duree + ' min)<br>' +
    esc(r.lieu) + '<br>' + esc(r.adresse) + '<br><br>' + esc(r.messageConfirmation) + '</div></div>';
  MailApp.sendEmail({ to: email, subject: 'Ton RDV chez ' + r.nom + ' — ' + quand, htmlBody: html, name: r.nom });
}
