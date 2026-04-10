const express  = require('express');
const http     = require('http');
const https    = require('https');
const { Server } = require('socket.io');
const fs       = require('fs');
const path     = require('path');
const os       = require('os');
const forge    = require('node-forge');
const session  = require('express-session');
const bcrypt   = require('bcryptjs');

const app = express();

const PORT        = 3443;
const PORT_HTTP   = 3000;
const PORT_TUNNEL = 3080;   // Port HTTP dédié au tunnel (données mobiles)
const DATA_FILE   = path.join(__dirname, 'data', 'pointages.json');
const AGENTS_FILE = path.join(__dirname, 'data', 'agents.json');
const USERS_FILE  = path.join(__dirname, 'data', 'users.json');

// ── Récupérer toutes les IPs locales non-internes ──
function getAllLocalIPs() {
  const interfaces = os.networkInterfaces();
  const ips = [];
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) ips.push(iface.address);
    }
  }
  return ips;
}

// Préférer l'IP WiFi, sinon la première disponible
function getLocalIP() {
  const interfaces = os.networkInterfaces();
  const wifiNames = ['wi-fi', 'wifi', 'wlan', 'wireless'];
  // Chercher d'abord un adaptateur WiFi
  for (const name of Object.keys(interfaces)) {
    if (wifiNames.some(w => name.toLowerCase().includes(w))) {
      for (const iface of interfaces[name]) {
        if (iface.family === 'IPv4' && !iface.internal) return iface.address;
      }
    }
  }
  // Sinon prendre la première IP non-interne
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) return iface.address;
    }
  }
  return '127.0.0.1';
}
const LOCAL_IP = getLocalIP();
const ALL_IPS  = getAllLocalIPs();

// ── Générer le certificat SSL avec SAN via node-forge ──
const CERT_FILE = path.join(__dirname, 'data', 'cert.json');
let sslCreds;
if (fs.existsSync(CERT_FILE)) {
  sslCreds = JSON.parse(fs.readFileSync(CERT_FILE, 'utf8'));
} else {
  console.log('  Génération du certificat SSL (première fois)...');

  // Générer la paire de clés RSA 2048
  const keys = forge.pki.rsa.generateKeyPair(2048);
  const cert  = forge.pki.createCertificate();

  cert.publicKey    = keys.publicKey;
  cert.serialNumber = '01';
  cert.validity.notBefore = new Date();
  cert.validity.notAfter  = new Date();
  cert.validity.notAfter.setFullYear(cert.validity.notBefore.getFullYear() + 10);

  const attrs = [
    { name: 'commonName',       value: 'localhost' },
    { name: 'organizationName', value: 'Gestion Qualite' },
    { name: 'countryName',      value: 'CI' }
  ];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);

  // Extensions obligatoires pour les navigateurs modernes
  cert.setExtensions([
    { name: 'basicConstraints', cA: true },
    { name: 'keyUsage',
      keyCertSign: true, digitalSignature: true,
      nonRepudiation: true, keyEncipherment: true, dataEncipherment: true },
    { name: 'extKeyUsage', serverAuth: true, clientAuth: true },
    { name: 'subjectAltName',
      altNames: [
        { type: 2, value: 'localhost' },
        { type: 7, ip: '127.0.0.1' },
        ...ALL_IPS.map(ip => ({ type: 7, ip }))
      ]
    }
  ]);

  cert.sign(keys.privateKey, forge.md.sha256.create());

  sslCreds = {
    key:  forge.pki.privateKeyToPem(keys.privateKey),
    cert: forge.pki.certificateToPem(cert)
  };
  fs.writeFileSync(CERT_FILE, JSON.stringify(sslCreds), 'utf8');
  console.log('  Certificat SSL généré avec succès (SHA-256 + SAN).');
}

const httpsServer = https.createServer({
  key:  sslCreds.key,
  cert: sslCreds.cert,
  minVersion: 'TLSv1.2'
}, app);

const httpServer = http.createServer((req, res) => {
  const host = (req.headers.host || '').split(':')[0];
  res.writeHead(301, { Location: `https://${host}:${PORT}${req.url}` });
  res.end();
});

const io = new Server(httpsServer);

// ── Helpers JSON ──
function lireJSON(file, defaut = []) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); }
  catch { return defaut; }
}
function ecrireJSON(file, data) {
  fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
}

// ── Store de sessions persistant (fichier JSON) ──
const SESSIONS_FILE = path.join(__dirname, 'data', 'sessions.json');
const session_store = require('express-session').Store;

class FileSessionStore extends session_store {
  constructor() {
    super();
    this._sessions = lireJSON(SESSIONS_FILE, {});
    // Nettoyage des sessions expirées toutes les heures
    setInterval(() => {
      const now = Date.now();
      let changed = false;
      for (const sid of Object.keys(this._sessions)) {
        const s = this._sessions[sid];
        if (s.cookie && s.cookie.expires && new Date(s.cookie.expires) < now) {
          delete this._sessions[sid];
          changed = true;
        }
      }
      if (changed) ecrireJSON(SESSIONS_FILE, this._sessions);
    }, 60 * 60 * 1000);
  }
  get(sid, cb) {
    const s = this._sessions[sid];
    cb(null, s || null);
  }
  set(sid, session, cb) {
    this._sessions[sid] = session;
    ecrireJSON(SESSIONS_FILE, this._sessions);
    cb(null);
  }
  destroy(sid, cb) {
    delete this._sessions[sid];
    ecrireJSON(SESSIONS_FILE, this._sessions);
    cb(null);
  }
}

const PHOTOS_DIR   = path.join(__dirname, 'public', 'photos');
const PIECES_DIR   = path.join(__dirname, 'public', 'pieces-jointes');
if (!fs.existsSync(PHOTOS_DIR))  fs.mkdirSync(PHOTOS_DIR,  { recursive: true });
if (!fs.existsSync(PIECES_DIR))  fs.mkdirSync(PIECES_DIR,  { recursive: true });

// ── Middleware ──
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// ── Sessions sécurisées et persistantes ──
app.use(session({
  store: new FileSessionStore(),
  secret: 'gestion-qualite-mafa-secret-2024',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: true,
    httpOnly: true,
    maxAge: 7 * 24 * 60 * 60 * 1000  // 7 jours
  }
}));

// ── Middleware de protection des routes admin ──
function requireAuth(req, res, next) {
  if (req.session && req.session.user) return next();
  res.redirect('/login');
}
function requireAdmin(req, res, next) {
  if (req.session && req.session.user && req.session.user.role === 'admin') return next();
  res.status(403).json({ erreur: 'Accès réservé à l\'administrateur.' });
}

app.use(express.static(path.join(__dirname, 'public')));

// ──────────────────────────────────────────
// ROUTES API
// ──────────────────────────────────────────

// Infos serveur (IP pour générer QR)
app.get('/api/infos', (req, res) => {
  res.json({ ip: LOCAL_IP, port: PORT, url: `https://${LOCAL_IP}:${PORT}/badge` });
});

// Liste agents
app.get('/api/agents', (req, res) => {
  res.json(lireJSON(AGENTS_FILE));
});

// Ajouter / modifier agent
app.post('/api/agents', (req, res) => {
  const { id, nom, prenom, poste, service } = req.body;
  if (!id || !nom || !prenom) return res.status(400).json({ erreur: 'Champs manquants' });
  const agents = lireJSON(AGENTS_FILE);
  const idx = agents.findIndex(a => a.id === id);
  const agent = { id: id.toUpperCase(), nom: nom.toUpperCase(), prenom, poste: poste || '', service: service || '' };
  if (idx >= 0) agents[idx] = agent; else agents.push(agent);
  ecrireJSON(AGENTS_FILE, agents);
  io.emit('agents-mis-a-jour', agents);
  res.json({ ok: true, agent });
});

// Supprimer agent
app.delete('/api/agents/:id', (req, res) => {
  const agents = lireJSON(AGENTS_FILE).filter(a => a.id !== req.params.id.toUpperCase());
  ecrireJSON(AGENTS_FILE, agents);
  io.emit('agents-mis-a-jour', agents);
  res.json({ ok: true });
});

// Tous les pointages
app.get('/api/pointages', (req, res) => {
  let pointages = lireJSON(DATA_FILE);
  const { date, agentId } = req.query;
  if (date)    pointages = pointages.filter(p => p.date === date);
  if (agentId) pointages = pointages.filter(p => p.agentId === agentId);
  res.json(pointages.reverse());
});

// Pointages du jour uniquement
app.get('/api/pointages/today', (req, res) => {
  const today = new Date().toLocaleDateString('fr-FR');
  const pointages = lireJSON(DATA_FILE).filter(p => p.date === today);
  res.json(pointages.reverse());
});

// Statistiques du jour
app.get('/api/stats', (req, res) => {
  const today = new Date().toLocaleDateString('fr-FR');
  const agents = lireJSON(AGENTS_FILE);
  const pointages = lireJSON(DATA_FILE);
  const duJour = pointages.filter(p => p.date === today);

  const presents = new Set(duJour.filter(p => p.type === 'entree').map(p => p.agentId));
  const retards  = duJour.filter(p => {
    if (p.type !== 'entree') return false;
    const [h, m] = p.heure.split(':').map(Number);
    return h > 9 || (h === 9 && m > 0);
  }).map(p => p.agentId);

  res.json({
    date: today,
    totalAgents: agents.length,
    presents: presents.size,
    absents: Math.max(0, agents.length - presents.size),
    retards: retards.length,
    derniersPointages: duJour.slice(-5).reverse()
  });
});

// ── BADGEAGE (soumis depuis le téléphone de l'agent) ──
app.post('/api/badger', (req, res) => {
  const { agentId, photo, motifRetard } = req.body;
  if (!agentId) return res.status(400).json({ erreur: 'ID agent manquant' });

  const agents = lireJSON(AGENTS_FILE);
  const agent = agents.find(a => a.id === agentId.toUpperCase());
  if (!agent) return res.status(404).json({ erreur: 'Agent non trouvé' });

  const pointages = lireJSON(DATA_FILE);
  const today = new Date().toLocaleDateString('fr-FR');
  const heure = new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
  const duJour = pointages.filter(p => p.agentId === agent.id && p.date === today);

  const aEntree = duJour.some(p => p.type === 'entree');
  const aSortie = duJour.some(p => p.type === 'sortie');

  if (aEntree && aSortie) {
    return res.json({ ok: false, message: 'Vous avez déjà badgé entrée et sortie aujourd\'hui.' });
  }

  const type = aEntree ? 'sortie' : 'entree';
  const [h, m] = heure.split(':').map(Number);
  const retard = type === 'entree' && (h > 9 || (h === 9 && m > 0));

  // ── Sauvegarder la photo ──
  const pointageId = Date.now().toString();
  let photoFichier = null;
  if (photo && photo.startsWith('data:image')) {
    try {
      const base64Data = photo.replace(/^data:image\/\w+;base64,/, '');
      const photoNom = `${pointageId}_${agent.id}_${type}.jpg`;
      fs.writeFileSync(path.join(PHOTOS_DIR, photoNom), Buffer.from(base64Data, 'base64'));
      photoFichier = `/photos/${photoNom}`;
    } catch (e) {
      console.error('Erreur sauvegarde photo:', e.message);
    }
  }

  const pointage = {
    id: pointageId,
    agentId: agent.id,
    nom: agent.nom,
    prenom: agent.prenom,
    nom_complet: agent.nom_complet || `${agent.prenom} ${agent.nom}`,
    poste: agent.poste,
    service: agent.service || '',
    date: today,
    heure,
    type,
    retard,
    motifRetard: (retard && motifRetard) ? motifRetard.trim() : undefined,
    photo: photoFichier,
    timestamp: new Date().toISOString()
  };

  pointages.push(pointage);
  ecrireJSON(DATA_FILE, pointages);

  // Notifier tous les admins en temps réel
  io.emit('nouveau-pointage', pointage);
  io.emit('stats-update');

  res.json({
    ok: true,
    type,
    retard,
    photo: photoFichier,
    agent: { nom: agent.nom, prenom: agent.prenom, nom_complet: agent.nom_complet, poste: agent.poste },
    heure
  });
});

// ── Supprimer un pointage (admin uniquement) ──
app.delete('/api/pointages/:id', requireAuth, requireAdmin, (req, res) => {
  const pointages = lireJSON(DATA_FILE).filter(p => p.id !== req.params.id);
  ecrireJSON(DATA_FILE, pointages);
  io.emit('stats-update');
  res.json({ ok: true });
});

// ─────────────────────────────────────────
// AUTHENTIFICATION
// ─────────────────────────────────────────

app.post('/api/login', (req, res) => {
  const { login, password } = req.body;
  const users = lireJSON(USERS_FILE);
  const user  = users.find(u => u.login === login && u.actif);
  if (!user || !bcrypt.compareSync(password, user.password))
    return res.status(401).json({ erreur: 'Identifiant ou mot de passe incorrect.' });
  req.session.user = { id: user.id, nom: user.nom, login: user.login, role: user.role };
  res.json({ ok: true, role: user.role, nom: user.nom });
});

app.post('/api/logout', requireAuth, (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

app.get('/api/me', requireAuth, (req, res) => {
  res.json(req.session.user);
});

app.post('/api/changer-mdp', requireAuth, (req, res) => {
  const { ancien, nouveau } = req.body;
  if (!nouveau || nouveau.length < 6)
    return res.status(400).json({ erreur: 'Minimum 6 caractères requis.' });
  const users = lireJSON(USERS_FILE);
  const idx   = users.findIndex(u => u.id === req.session.user.id);
  if (idx < 0 || !bcrypt.compareSync(ancien, users[idx].password))
    return res.status(401).json({ erreur: 'Ancien mot de passe incorrect.' });
  users[idx].password = bcrypt.hashSync(nouveau, 10);
  ecrireJSON(USERS_FILE, users);
  res.json({ ok: true });
});

app.get('/api/users', requireAuth, requireAdmin, (req, res) => {
  res.json(lireJSON(USERS_FILE).map(u => ({ ...u, password: undefined })));
});

app.post('/api/users', requireAuth, requireAdmin, (req, res) => {
  const { nom, login, password, role } = req.body;
  if (!login || !password || !nom)
    return res.status(400).json({ erreur: 'Champs manquants.' });
  const users = lireJSON(USERS_FILE);
  if (users.find(u => u.login === login))
    return res.status(400).json({ erreur: 'Ce login existe déjà.' });
  users.push({ id: Date.now().toString(), nom, login,
    password: bcrypt.hashSync(password, 10), role: role || 'superviseur', actif: true });
  ecrireJSON(USERS_FILE, users);
  res.json({ ok: true });
});

app.delete('/api/users/:id', requireAuth, requireAdmin, (req, res) => {
  if (req.session.user.id === req.params.id)
    return res.status(400).json({ erreur: 'Impossible de supprimer votre propre compte.' });
  ecrireJSON(USERS_FILE, lireJSON(USERS_FILE).filter(u => u.id !== req.params.id));
  res.json({ ok: true });
});

// ─────────────────────────────────────────
// DEMANDES (permissions / absences)
// ─────────────────────────────────────────
const DEMANDES_FILE = path.join(__dirname, 'data', 'demandes.json');

// Soumettre une demande (public — agents sans compte)
app.post('/api/demandes', (req, res) => {
  const { agentId, nom, service, type, dateDebut, dateFin, heure, motif, tel } = req.body;
  if (!agentId || !nom || !type || !dateDebut || !motif)
    return res.status(400).json({ erreur: 'Champs obligatoires manquants.' });

  const demandes = lireJSON(DEMANDES_FILE);
  const demandeId = Date.now().toString();

  // ── Sauvegarder les pièces jointes ──
  const { piecesJointes } = req.body;
  const fichiersSauves = [];
  if (Array.isArray(piecesJointes)) {
    for (const pj of piecesJointes) {
      try {
        const ext = pj.nom.split('.').pop().replace(/[^a-z0-9]/gi, '').toLowerCase() || 'bin';
        const nomFichier = `${demandeId}_${fichiersSauves.length + 1}.${ext}`;
        const base64Data = pj.data.replace(/^data:[^;]+;base64,/, '');
        fs.writeFileSync(path.join(PIECES_DIR, nomFichier), Buffer.from(base64Data, 'base64'));
        fichiersSauves.push({ nom: pj.nom, chemin: `/pieces-jointes/${nomFichier}` });
      } catch (e) { console.error('Erreur pièce jointe:', e.message); }
    }
  }

  const demande = {
    id: demandeId,
    agentId: agentId.toUpperCase(),
    nom, service, type, dateDebut, dateFin: dateFin || null,
    heure: heure || null, motif, tel: tel || '',
    piecesJointes: fichiersSauves,
    statut: 'en_attente',
    commentaire: '',
    createdAt: new Date().toISOString()
  };
  demandes.push(demande);
  ecrireJSON(DEMANDES_FILE, demandes);

  // Notifier les admins en temps réel
  io.emit('nouvelle-demande', demande);
  res.json({ ok: true, demande });
});

// Liste toutes les demandes (admin)
app.get('/api/demandes', requireAuth, (req, res) => {
  let demandes = lireJSON(DEMANDES_FILE);
  if (req.query.statut) demandes = demandes.filter(d => d.statut === req.query.statut);
  res.json(demandes.sort((a, b) => b.createdAt.localeCompare(a.createdAt)));
});

// Changer le statut d'une demande (admin)
app.patch('/api/demandes/:id', requireAuth, requireAdmin, (req, res) => {
  const { statut, commentaire } = req.body;
  const demandes = lireJSON(DEMANDES_FILE);
  const idx = demandes.findIndex(d => d.id === req.params.id);
  if (idx < 0) return res.status(404).json({ erreur: 'Demande introuvable.' });
  demandes[idx].statut = statut || demandes[idx].statut;
  demandes[idx].commentaire = commentaire !== undefined ? commentaire : demandes[idx].commentaire;
  demandes[idx].traitePar = req.session.user.nom;
  demandes[idx].traiteAt = new Date().toISOString();
  ecrireJSON(DEMANDES_FILE, demandes);
  io.emit('demande-mise-a-jour', demandes[idx]);
  res.json({ ok: true, demande: demandes[idx] });
});

// Supprimer une demande (admin)
app.delete('/api/demandes/:id', requireAuth, requireAdmin, (req, res) => {
  ecrireJSON(DEMANDES_FILE, lireJSON(DEMANDES_FILE).filter(d => d.id !== req.params.id));
  res.json({ ok: true });
});

// ─────────────────────────────────────────
// PAGES
// ─────────────────────────────────────────
app.get('/login',     (_, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/badge',     (_, res) => res.sendFile(path.join(__dirname, 'public', 'badge.html')));
app.get('/affichage', (_, res) => res.sendFile(path.join(__dirname, 'public', 'affichage.html')));
app.get('/demandes',          (_, res) => res.sendFile(path.join(__dirname, 'public', 'demandes.html')));
app.get('/affichage-mobile',  (_, res) => res.sendFile(path.join(__dirname, 'public', 'affichage-mobile.html')));
app.get('/admin',     requireAuth, (_, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));
app.get('/', (req, res) => req.session && req.session.user ? res.redirect('/admin') : res.redirect('/login'));

// ─────────────────────────────────────────
// SERVEUR TUNNEL HTTP (port 3080)
// App publique minimale — uniquement pour les agents via données mobiles
// ─────────────────────────────────────────
const appPublic = express();
appPublic.use(express.json({ limit: '10mb' }));
appPublic.use(express.static(path.join(__dirname, 'public')));

// Pages accessibles via tunnel
appPublic.get('/',                  (_, res) => res.sendFile(path.join(__dirname, 'public', 'demandes.html')));
appPublic.get('/demandes',          (_, res) => res.sendFile(path.join(__dirname, 'public', 'demandes.html')));
appPublic.get('/badge',             (_, res) => res.sendFile(path.join(__dirname, 'public', 'badge.html')));
appPublic.get('/superviseur',       (_, res) => res.sendFile(path.join(__dirname, 'public', 'superviseur.html')));
appPublic.get('/affichage-mobile',  (_, res) => res.sendFile(path.join(__dirname, 'public', 'affichage-mobile.html')));

// API superviseur via tunnel
appPublic.get('/api/superviseur/stats', (req, res) => {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token !== SUPERVISEUR_TOKEN) return res.status(403).json({ erreur: 'Accès refusé.' });
  const today = new Date().toLocaleDateString('fr-FR');
  const agents = lireJSON(AGENTS_FILE);
  const pointages = lireJSON(DATA_FILE);
  const duJour = pointages.filter(p => p.date === today);
  const presents = new Set(duJour.filter(p => p.type === 'entree').map(p => p.agentId));
  const retards = duJour.filter(p => { if (p.type !== 'entree') return false; const [h,m]=p.heure.split(':').map(Number); return h>9||(h===9&&m>0); }).length;
  res.json({ totalAgents: agents.length, presents: presents.size, absents: Math.max(0, agents.length - presents.size), retards });
});
appPublic.get('/api/superviseur/agents', (req, res) => {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token !== SUPERVISEUR_TOKEN) return res.status(403).json({ erreur: 'Accès refusé.' });
  res.json(lireJSON(AGENTS_FILE));
});
appPublic.get('/api/superviseur/pointages-today', (req, res) => {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token !== SUPERVISEUR_TOKEN) return res.status(403).json({ erreur: 'Accès refusé.' });
  const today = new Date().toLocaleDateString('fr-FR');
  res.json(lireJSON(DATA_FILE).filter(p => p.date === today).reverse());
});
appPublic.get('/api/superviseur/pointages', (req, res) => {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token !== SUPERVISEUR_TOKEN) return res.status(403).json({ erreur: 'Accès refusé.' });
  let pts = lireJSON(DATA_FILE);
  if (req.query.date) pts = pts.filter(p => p.date === req.query.date);
  res.json(pts.reverse());
});
appPublic.get('/api/superviseur/demandes', (req, res) => {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token !== SUPERVISEUR_TOKEN) return res.status(403).json({ erreur: 'Accès refusé.' });
  let dem = lireJSON(DEMANDES_FILE);
  if (req.query.statut)  dem = dem.filter(d => d.statut  === req.query.statut);
  if (req.query.service) dem = dem.filter(d => d.service === req.query.service);
  res.json(dem.sort((a,b) => b.createdAt.localeCompare(a.createdAt)));
});
appPublic.get('/api/tunnel-info', (_, res) => {
  const data = fs.existsSync(TUNNEL_FILE) ? lireJSON(TUNNEL_FILE, {}) : {};
  res.json({ badge: data.url ? `${data.url}/badge` : null, demandes: data.url ? `${data.url}/demandes` : null, actif: !!data.url });
});

// API agents (lecture seule — pour le badge)
appPublic.get('/api/agents', (_, res) => res.json(lireJSON(AGENTS_FILE)));

// API badgeage via tunnel
appPublic.post('/api/badger', (req, res) => {
  const { agentId, photo, motifRetard } = req.body;
  if (!agentId) return res.status(400).json({ erreur: 'ID agent manquant' });
  const agents = lireJSON(AGENTS_FILE);
  const agent = agents.find(a => a.id === agentId.toUpperCase());
  if (!agent) return res.status(404).json({ erreur: 'Agent non trouvé' });
  const pointages = lireJSON(DATA_FILE);
  const today = new Date().toLocaleDateString('fr-FR');
  const heure = new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
  const duJour = pointages.filter(p => p.agentId === agent.id && p.date === today);
  const aEntree = duJour.some(p => p.type === 'entree');
  const aSortie = duJour.some(p => p.type === 'sortie');
  if (aEntree && aSortie)
    return res.json({ ok: false, message: 'Vous avez déjà badgé entrée et sortie aujourd\'hui.' });
  const type = aEntree ? 'sortie' : 'entree';
  const [h, m] = heure.split(':').map(Number);
  const retard = type === 'entree' && (h > 9 || (h === 9 && m > 0));
  const pointageId = Date.now().toString();
  let photoFichier = null;
  if (photo && photo.startsWith('data:image')) {
    try {
      const base64Data = photo.replace(/^data:image\/\w+;base64,/, '');
      const photoNom = `${pointageId}_${agent.id}_${type}.jpg`;
      fs.writeFileSync(path.join(PHOTOS_DIR, photoNom), Buffer.from(base64Data, 'base64'));
      photoFichier = `/photos/${photoNom}`;
    } catch (e) { console.error('Erreur photo:', e.message); }
  }
  const pointage = {
    id: pointageId, agentId: agent.id, nom: agent.nom, prenom: agent.prenom,
    nom_complet: agent.nom_complet || `${agent.prenom} ${agent.nom}`,
    poste: agent.poste, service: agent.service || '',
    date: today, heure, type, retard,
    motifRetard: (retard && motifRetard) ? motifRetard.trim() : undefined,
    photo: photoFichier,
    source: 'mobile-data', timestamp: new Date().toISOString()
  };
  pointages.push(pointage);
  ecrireJSON(DATA_FILE, pointages);
  io.emit('nouveau-pointage', pointage);
  io.emit('stats-update');
  res.json({ ok: true, type, retard, photo: photoFichier,
    agent: { nom: agent.nom, prenom: agent.prenom, nom_complet: agent.nom_complet, poste: agent.poste }, heure });
});

// Route POST demandes (tunnel) — réutilise la même logique
appPublic.post('/api/demandes', (req, res) => {
  // Rediriger vers le handler principal en ajoutant source mobile
  req.body.source = 'mobile';
  // Appeler directement la logique
  const { agentId, nom, service, type, dateDebut, dateFin, heure, motif, tel, piecesJointes } = req.body;
  if (!agentId || !nom || !type || !dateDebut || !motif)
    return res.status(400).json({ erreur: 'Champs obligatoires manquants.' });
  const demandes = lireJSON(DEMANDES_FILE);
  const demandeId = Date.now().toString();
  const fichiersSauves = [];
  if (Array.isArray(piecesJointes)) {
    for (const pj of piecesJointes) {
      try {
        const ext = pj.nom.split('.').pop().replace(/[^a-z0-9]/gi, '').toLowerCase() || 'bin';
        const nomFichier = `${demandeId}_${fichiersSauves.length + 1}.${ext}`;
        const base64Data = pj.data.replace(/^data:[^;]+;base64,/, '');
        fs.writeFileSync(path.join(PIECES_DIR, nomFichier), Buffer.from(base64Data, 'base64'));
        fichiersSauves.push({ nom: pj.nom, chemin: `/pieces-jointes/${nomFichier}` });
      } catch (e) { console.error('Erreur pièce jointe (tunnel):', e.message); }
    }
  }
  const demande = {
    id: demandeId, agentId: agentId.toUpperCase(),
    nom, service, type, dateDebut, dateFin: dateFin || null,
    heure: heure || null, motif, tel: tel || '',
    piecesJointes: fichiersSauves,
    statut: 'en_attente', commentaire: '',
    source: 'mobile', createdAt: new Date().toISOString()
  };
  demandes.push(demande);
  ecrireJSON(DEMANDES_FILE, demandes);
  io.emit('nouvelle-demande', demande);
  res.json({ ok: true, demande });
});

const tunnelHttpServer = http.createServer(appPublic);

// ── Fichier URL tunnel ──
const TUNNEL_FILE = path.join(__dirname, 'data', 'tunnel.json');
let tunnelUrl = null;

function sauvegarderTunnel(url) {
  tunnelUrl = url;
  ecrireJSON(TUNNEL_FILE, { url, updatedAt: new Date().toISOString() });
  io.emit('tunnel-url-update', { url });
}

// ── Tunnel Cloudflare (cloudflared) — fiable, gratuit, sans compte ──
const { spawn } = require('child_process');
const CLOUDFLARED = path.join(__dirname, 'cloudflared.exe');
let _cfProcess = null;
let _tunnelActif = false;

function lancerTunnel() {
  if (_tunnelActif) return;
  if (!fs.existsSync(CLOUDFLARED)) {
    console.error('  cloudflared.exe introuvable — tunnel désactivé');
    return;
  }
  _tunnelActif = true;
  console.log('  Démarrage du tunnel Cloudflare...');

  _cfProcess = spawn(CLOUDFLARED, [
    'tunnel', '--url', `http://localhost:${PORT_TUNNEL}`,
    '--no-autoupdate'
  ], { windowsHide: true });

  // Cloudflare écrit l'URL dans stderr
  _cfProcess.stderr.on('data', (data) => {
    const txt = data.toString();
    const match = txt.match(/https:\/\/[a-z0-9\-]+\.trycloudflare\.com/);
    if (match && match[0] !== tunnelUrl) {
      sauvegarderTunnel(match[0]);
      console.log(`\n  📱 Lien données mobiles : ${match[0]}/demandes\n`);
    }
  });

  _cfProcess.on('exit', (code) => {
    console.log(`  ⚠️  Tunnel Cloudflare arrêté (code ${code}) — relance dans 10s...`);
    _cfProcess = null;
    _tunnelActif = false;
    sauvegarderTunnel(null);
    setTimeout(lancerTunnel, 10000);
  });

  _cfProcess.on('error', (err) => {
    console.error('  Tunnel erreur:', err.message);
    _cfProcess = null;
    _tunnelActif = false;
    sauvegarderTunnel(null);
    setTimeout(lancerTunnel, 10000);
  });
}

// Vérification toutes les 2 minutes — relance si processus mort
setInterval(() => {
  if (!_cfProcess || !tunnelUrl) {
    _tunnelActif = false;
    lancerTunnel();
  }
}, 2 * 60 * 1000);

// ── API : URL publique courante (admin) ──
app.get('/api/public-url', requireAuth, (req, res) => {
  const data = fs.existsSync(TUNNEL_FILE) ? lireJSON(TUNNEL_FILE, {}) : {};
  res.json({
    urlLocale:   `https://${LOCAL_IP}:${PORT}/demandes`,
    urlPublique: data.url ? `${data.url}/demandes` : null,
    tunnel: !!data.url
  });
});

// ─────────────────────────────────────────
// ROUTES SUPERVISEUR (token URL, lecture seule)
// ─────────────────────────────────────────
const SUPERVISEUR_TOKEN = '613f01aaac557e34545efc1c73c5337e';

function requireSuperviseur(req, res, next) {
  const token = req.query.token || req.headers['x-superviseur-token'];
  if (token === SUPERVISEUR_TOKEN) return next();
  res.status(403).json({ erreur: 'Accès refusé.' });
}

app.get('/api/superviseur/stats',           requireSuperviseur, (req, res) => {
  const today    = new Date().toLocaleDateString('fr-FR');
  const agents   = lireJSON(AGENTS_FILE);
  const pointages= lireJSON(DATA_FILE);
  const duJour   = pointages.filter(p => p.date === today);
  const presents = new Set(duJour.filter(p => p.type === 'entree').map(p => p.agentId));
  const retards  = duJour.filter(p => { if (p.type !== 'entree') return false; const [h,m]=p.heure.split(':').map(Number); return h>9||(h===9&&m>0); }).length;
  res.json({ totalAgents:agents.length, presents:presents.size, absents:Math.max(0,agents.length-presents.size), retards });
});

app.get('/api/superviseur/agents',          requireSuperviseur, (req, res) => res.json(lireJSON(AGENTS_FILE)));

app.get('/api/superviseur/pointages-today', requireSuperviseur, (req, res) => {
  const today = new Date().toLocaleDateString('fr-FR');
  res.json(lireJSON(DATA_FILE).filter(p => p.date === today).reverse());
});

app.get('/api/superviseur/pointages',       requireSuperviseur, (req, res) => {
  let pts = lireJSON(DATA_FILE);
  if (req.query.date) pts = pts.filter(p => p.date === req.query.date);
  res.json(pts.reverse());
});

app.get('/api/superviseur/demandes',        requireSuperviseur, (req, res) => {
  let dem = lireJSON(DEMANDES_FILE);
  if (req.query.statut)  dem = dem.filter(d => d.statut  === req.query.statut);
  if (req.query.service) dem = dem.filter(d => d.service === req.query.service);
  res.json(dem.sort((a,b) => b.createdAt.localeCompare(a.createdAt)));
});

app.get('/superviseur', (_, res) => res.sendFile(path.join(__dirname, 'public', 'superviseur.html')));

// ── API : URL tunnel publique (sans auth — pour les pages d'affichage) ──
app.get('/api/tunnel-info', (req, res) => {
  const data = fs.existsSync(TUNNEL_FILE) ? lireJSON(TUNNEL_FILE, {}) : {};
  res.json({
    badge:    data.url ? `${data.url}/badge`    : null,
    demandes: data.url ? `${data.url}/demandes` : null,
    actif:    !!data.url
  });
});

// ── WebSocket ──
io.on('connection', (socket) => {
  socket.on('disconnect', () => {});
  // Envoyer l'URL tunnel dès la connexion
  if (tunnelUrl) socket.emit('tunnel-url-update', { url: tunnelUrl });
});

// ── Démarrage séquentiel ──
httpServer.listen(PORT_HTTP, '0.0.0.0', () => {
  console.log(`  Redirection HTTP  : http://localhost:${PORT_HTTP} → HTTPS`);
}).on('error', (e) => {
  if (e.code === 'EADDRINUSE')
    console.log(`  (Port ${PORT_HTTP} déjà utilisé — redirection HTTP ignorée)`);
});

tunnelHttpServer.listen(PORT_TUNNEL, '0.0.0.0', () => {
  console.log(`  Serveur tunnel    : http://0.0.0.0:${PORT_TUNNEL} (données mobiles)`);
}).on('error', (e) => {
  if (e.code === 'EADDRINUSE') console.log(`  (Port ${PORT_TUNNEL} déjà utilisé)`);
});

httpsServer.listen(PORT, '0.0.0.0', () => {
  console.log('\n╔══════════════════════════════════════════════════════╗');
  console.log('║     SYSTÈME GESTION QUALITÉ — SERVEUR ACTIF (HTTPS) ║');
  console.log('╠══════════════════════════════════════════════════════╣');
  console.log(`║  Dashboard Admin : https://localhost:${PORT}/admin        ║`);
  console.log(`║  Badgeage agents : https://${LOCAL_IP}:${PORT}/badge  ║`);
  console.log(`║  Demandes (WiFi) : https://${LOCAL_IP}:${PORT}/demandes║`);
  console.log(`║  Affichage QR    : https://localhost:${PORT}/affichage    ║`);
  console.log('╠══════════════════════════════════════════════════════╣');
  console.log('║  Tunnel en cours de démarrage (données mobiles)...  ║');
  console.log('╚══════════════════════════════════════════════════════╝\n');

  // Lancer le tunnel après démarrage du serveur
  lancerTunnel();
});
