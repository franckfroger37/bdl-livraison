import React, { useState, useEffect } from 'react';
import { MapPin, Clock, Package, FileText, Download, ArrowRight, Home, CheckCircle, Settings } from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// ─── Configuration ───────────────────────────────────────────────────────────
const CONFIG_DEFAULT = {
  cheminExcel: '',
  email1: '',
  email2: '',
  email3: '',
  adresseUnite: '',
  gpsUnite: '',
  emailjsServiceId: '',
  emailjsTemplateId: '',
  emailjsPublicKey: '',
};

async function chargerConfigAsync() {
  let baseConfig = { ...CONFIG_DEFAULT };
  
  // 1. Charger config.json depuis public/
  try {
    const response = await fetch('/config.json?t=' + Date.now());
    if (response.ok) {
      const fileConfig = await response.json();
      baseConfig = { ...baseConfig, ...fileConfig };
    }
  } catch {}
  
  // 2. Surcharger avec localStorage (modifs utilisateur)
  try {
    const localRaw = localStorage.getItem('bdl-config');
    if (localRaw) {
      const localConfig = JSON.parse(localRaw);
      baseConfig = { ...baseConfig, ...localConfig };
    }
  } catch {}
  
  return baseConfig;
}

function chargerConfig() {
  try {
    const raw = localStorage.getItem('bdl-config');
    return raw ? { ...CONFIG_DEFAULT, ...JSON.parse(raw) } : { ...CONFIG_DEFAULT };
  } catch { return { ...CONFIG_DEFAULT }; }
}

function sauverConfig(cfg) {
  localStorage.setItem('bdl-config', JSON.stringify(cfg));
}

// ─── Compression photo ───────────────────────────────────────────────────────
function compresserPhoto(dataUrl, maxWidth = 800) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let w = img.width, h = img.height;
      if (w > maxWidth) { h = Math.round(h * maxWidth / w); w = maxWidth; }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.6));
    };
    img.src = dataUrl;
  });
}

// ─── Écran de configuration ──────────────────────────────────────────────────
function VueConfig({ config, onSave, onRetour }) {
  const [cfg, setCfg] = useState({ ...CONFIG_DEFAULT, ...config });
  const refImport = React.useRef(null);

  // Synchroniser avec config.json à l'ouverture
  useEffect(() => {
    chargerConfigAsync().then(cfgFichier => {
      setCfg(prev => ({ ...prev, ...cfgFichier }));
    });
  }, []);

  const champ = (label, key, placeholder, type = 'text') => (
    <div style={{ marginBottom:'16px' }}>
      <label style={{ display:'block', fontWeight:'600', color:'#374151', marginBottom:'6px', fontSize:'15px' }}>{label}</label>
      <input
        type={type}
        value={cfg[key]}
        onChange={e => setCfg(prev => ({ ...prev, [key]: e.target.value }))}
        placeholder={placeholder}
        style={{ width:'100%', padding:'14px', fontSize:'15px', border:'2px solid #d1d5db', borderRadius:'10px', outline:'none', boxSizing:'border-box' }}
      />
    </div>
  );

  const sauver = () => {
    sauverConfig(cfg);
    if (onSave) onSave(cfg);
    alert("Configuration sauvegardée ✅");
    onRetour();
  };

  // Exporter la config comme fichier JSON téléchargeable
  const exporterConfig = () => {
    const blob = new Blob([JSON.stringify(cfg, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bdl-config.json';
    a.click();
    URL.revokeObjectURL(url);
  };

  // Importer une config depuis un fichier JSON
  const importerConfig = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        const merged = { ...CONFIG_DEFAULT, ...parsed };
        setCfg(merged);
        sauverConfig(merged);
        if (onSave) onSave(merged);
        alert("Configuration importée avec succès ✅");
      } catch {
        alert("Fichier invalide. Vérifiez que c'est bien un fichier bdl-config.json");
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'600px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'32px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'26px', fontWeight:'bold', marginBottom:'8px' }}>
          <Settings size={28} style={{ verticalAlign:'middle', marginRight:'10px', color:'#6b7280' }} />
          Configuration
        </h2>
        <p style={{ color:'#6b7280', fontSize:'14px', marginBottom:'24px' }}>
          Les paramètres sont sauvegardés sur cet appareil. Utilisez Export/Import pour les transférer.
        </p>

        {/* Import / Export */}
        <div style={{ background:'#f8fafc', border:'2px solid #e2e8f0', borderRadius:'12px', padding:'16px', marginBottom:'24px' }}>
          <p style={{ fontWeight:'700', color:'#374151', marginBottom:'12px', fontSize:'15px' }}>🔄 Transfert de configuration</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'10px' }}>
            <button onClick={exporterConfig}
              style={{ padding:'14px', background:'#0369a1', color:'white', border:'none', borderRadius:'10px', fontSize:'14px', fontWeight:'bold', cursor:'pointer' }}>
              📤 Exporter
            </button>
            <button onClick={() => refImport.current.click()}
              style={{ padding:'14px', background:'#7c3aed', color:'white', border:'none', borderRadius:'10px', fontSize:'14px', fontWeight:'bold', cursor:'pointer' }}>
              📥 Importer
            </button>
            <input ref={refImport} type="file" accept=".json" onChange={importerConfig}
              style={{ position:'absolute', width:'1px', height:'1px', opacity:0, pointerEvents:'none' }} />
          </div>
          <p style={{ fontSize:'12px', color:'#94a3b8', margin:'10px 0 0', textAlign:'center' }}>
            Exportez depuis le PC, importez sur le mobile
          </p>
        </div>

        <div style={{ background:'#f0f9ff', border:'2px solid #bae6fd', borderRadius:'12px', padding:'16px', marginBottom:'20px' }}>
          <p style={{ fontWeight:'700', color:'#0369a1', marginBottom:'12px', fontSize:'15px' }}>📁 Fichiers Excel des tournées</p>
          {champ('Chemin ou lien', 'cheminExcel', 'Ex: \\\\serveur\\tournees\\ ou https://...')}
        </div>

        <div style={{ background:'#fdf4ff', border:'2px solid #e9d5ff', borderRadius:'12px', padding:'16px', marginBottom:'20px' }}>
          <p style={{ fontWeight:'700', color:'#7c3aed', marginBottom:'12px', fontSize:'15px' }}>📧 Envoi du rapport PDF</p>
          {champ('Email 1', 'email1', 'responsable@bdl.fr', 'email')}
          {champ('Email 2 (optionnel)', 'email2', 'chef@bdl.fr', 'email')}
          {champ('Email 3 (optionnel)', 'email3', 'archive@bdl.fr', 'email')}
        </div>

        <div style={{ background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'12px', padding:'16px', marginBottom:'20px' }}>
          <p style={{ fontWeight:'700', color:'#92400e', marginBottom:'6px', fontSize:'15px' }}>📎 Envoi avec pièce jointe (EmailJS)</p>
          <p style={{ fontSize:'12px', color:'#b45309', marginBottom:'12px' }}>
            Optionnel — Créez un compte sur <strong>emailjs.com</strong> pour envoyer le PDF en pièce jointe.
            Sans ça, le bouton télécharge le PDF et ouvre votre client mail.
          </p>
          {champ('Service ID', 'emailjsServiceId', 'Ex: service_xxxxxxx')}
          {champ('Template ID', 'emailjsTemplateId', 'Ex: template_xxxxxxx')}
          {champ('Clé publique', 'emailjsPublicKey', 'Ex: xxxxxxxxxxxxxx')}
        </div>

        <div style={{ background:'#f0fdf4', border:'2px solid #86efac', borderRadius:'12px', padding:'16px', marginBottom:'28px' }}>
          <p style={{ fontWeight:'700', color:'#15803d', marginBottom:'12px', fontSize:'15px' }}>🏢 Unité de départ / retour</p>
          {champ("Adresse de l'unité", 'adresseUnite', 'Ex: 12 Rue de la Blanchisserie, 37000 Tours')}
          {champ('Coordonnées GPS (optionnel)', 'gpsUnite', 'Ex: 47.3941, 0.6848')}
        </div>

        <div style={{ display:'flex', gap:'12px' }}>
          <button onClick={onRetour}
            style={{ flex:1, padding:'18px', background:'#e5e7eb', color:'#374151', border:'none', borderRadius:'12px', fontSize:'17px', fontWeight:'bold', cursor:'pointer' }}>
            Annuler
          </button>
          <button onClick={sauver}
            style={{ flex:2, padding:'18px', background:'#16a34a', color:'white', border:'none', borderRadius:'12px', fontSize:'17px', fontWeight:'bold', cursor:'pointer' }}>
            ✅ Sauvegarder
          </button>
        </div>
      </div>
    </div>
  );
}
// ─── Écran de connexion ──────────────────────────────────────────────────────
function VueLogin({ onLogin, onConfig }) {
  const [nom, setNom] = useState('');
  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#1d4ed8,#1e40af)', display:'flex', alignItems:'center', justifyContent:'center', padding:'16px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'40px', width:'100%', maxWidth:'420px', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign:'center', marginBottom:'32px' }}>
          <Package size={72} style={{ color:'#2563eb', margin:'0 auto 16px' }} />
          <h1 style={{ fontSize:'32px', fontWeight:'bold', color:'#1f2937', margin:0 }}>BDL-LIVRAISON</h1>
          <p style={{ color:'#6b7280', marginTop:'8px' }}>Gestion de tournées</p>
        </div>
        <div style={{ marginBottom:'24px' }}>
          <label style={{ display:'block', fontSize:'18px', fontWeight:'600', color:'#374151', marginBottom:'10px' }}>Nom du chauffeur</label>
          <input
            type="text" value={nom}
            onChange={e => setNom(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && nom.trim() && onLogin(nom.trim())}
            placeholder="Votre nom"
            style={{ width:'100%', padding:'16px', fontSize:'20px', border:'2px solid #d1d5db', borderRadius:'12px', outline:'none', boxSizing:'border-box' }}
          />
        </div>
        <button onClick={() => nom.trim() && onLogin(nom.trim())}
          style={{ width:'100%', padding:'20px', background:'#2563eb', color:'white', border:'none', borderRadius:'12px', fontSize:'20px', fontWeight:'bold', cursor:'pointer', marginBottom:'12px' }}>
          DÉMARRER
        </button>
        <button onClick={onConfig}
          style={{ width:'100%', padding:'14px', background:'#f3f4f6', color:'#6b7280', border:'2px solid #e5e7eb', borderRadius:'12px', fontSize:'15px', fontWeight:'600', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px' }}>
          <Settings size={18} /> Configuration
        </button>
      </div>
    </div>
  );
}

// ─── Écran d'import Excel ────────────────────────────────────────────────────
function VueImport({ onImport, onConfig }) {
  const cfg = chargerConfig();
  const [tourneesPublic, setTourneesPublic] = useState([]);

  // Charger la liste des tournées depuis public/tournees/
  // Se relance à chaque fois que le composant s'affiche
  useEffect(() => {
    console.log('🔄 Chargement de la liste des tournées...');
    fetch('/tournees/liste.json?t=' + Date.now())
      .then(res => {
        console.log('Réponse liste.json:', res.status, res.ok);
        if (!res.ok) {
          console.warn('liste.json non trouvé - créez public/tournees/liste.json');
          return [];
        }
        return res.json();
      })
      .then(fichiers => {
        console.log('Tournées chargées:', fichiers);
        setTourneesPublic(fichiers || []);
      })
      .catch(err => {
        console.error('Erreur chargement liste.json:', err);
        setTourneesPublic([]);
      });
  }, []); // Se charge au montage du composant

  const chargerDepuisPublic = async (nomFichier) => {
    try {
      const response = await fetch(`/tournees/${nomFichier}?t=` + Date.now());
      if (!response.ok) {
        alert(`Impossible de charger ${nomFichier}. Vérifiez qu'il existe dans public/tournees/`);
        return;
      }
      const arrayBuffer = await response.arrayBuffer();
      
      // Traiter directement l'ArrayBuffer avec XLSX
      try {
        const data = new Uint8Array(arrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const clients = [];
        let current = null;
        rows.forEach(row => {
          if (row.Client && String(row.Client).trim()) {
            if (current) clients.push(current);
            current = {
              id: `C${clients.length + 1}`,
              nom: String(row.Client).trim(),
              adresse: row.Adresse || '',
              code1: row.Code1 || '', code2: row.Code2 || '', code3: row.Code3 || '',
              telephone: row.Telephone || '',
              commentaireFixe: row.Commentaire || '',
              services: []
            };
          }
          if (row.Service && current) {
            current.services.push({
              id: `S${current.services.length + 1}`,
              nom: String(row.Service).trim(),
              cabrisPrevu: parseInt(row.CabrisPrevu) || 0,
              cabrisRecuperes: 0
            });
          }
        });
        if (current) clients.push(current);
        if (clients.length === 0) { 
          alert("Aucun client trouvé dans " + nomFichier); 
          return; 
        }
        onImport({ 
          id: `T-${new Date().toISOString().split('T')[0]}`, 
          date: new Date().toISOString(), 
          clients 
        });
      } catch (parseErr) {
        alert("Erreur de lecture du fichier Excel. Format invalide.");
        console.error(parseErr);
      }
    } catch (err) {
      alert(`Erreur de chargement de ${nomFichier}`);
      console.error(err);
    }
  };

  const traiterFichierExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const clients = [];
        let current = null;
        rows.forEach(row => {
          if (row.Client && String(row.Client).trim()) {
            if (current) clients.push(current);
            current = {
              id: `C${clients.length + 1}`,
              nom: String(row.Client).trim(),
              adresse: row.Adresse || '',
              code1: row.Code1 || '', code2: row.Code2 || '', code3: row.Code3 || '',
              telephone: row.Telephone || '',
              commentaireFixe: row.Commentaire || '',
              services: []
            };
          }
          if (row.Service && current) {
            current.services.push({
              id: `S${current.services.length + 1}`,
              nom: String(row.Service).trim(),
              cabrisPrevu: parseInt(row.CabrisPrevu) || 0,
              cabrisRecuperes: 0
            });
          }
        });
        if (current) clients.push(current);
        if (clients.length === 0) { alert("Aucun client trouvé. Vérifiez le format du fichier Excel."); return; }
        onImport({ id: `T-${new Date().toISOString().split('T')[0]}`, date: new Date().toISOString(), clients });
      } catch { alert("Erreur de lecture du fichier Excel. Vérifiez le format."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (file) traiterFichierExcel(file);
  };

    return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'600px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'40px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'16px' }}>
          <h2 style={{ fontSize:'20px', fontWeight:'bold', color:'#1f2937', margin:0 }}>📥 Choisir une tournée</h2>
          <button 
            onClick={() => {
              fetch('/tournees/liste.json?t=' + Date.now())
                .then(r => r.json())
                .then(data => setTourneesPublic(data))
                .catch(() => {});
            }}
            style={{ background:'#dbeafe', border:'2px solid #93c5fd', borderRadius:'10px', padding:'8px 12px', cursor:'pointer', color:'#1e40af', fontSize:'14px', fontWeight:'600' }}>
            🔄 Actualiser
          </button>
        </div>
        {/* Debug - État actuel */}
        <div style={{ marginBottom:'16px', padding:'12px', background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'10px', fontSize:'13px' }}>
          <strong>📊 Debug:</strong> {tourneesPublic.length} tournée(s) chargée(s)
          {tourneesPublic.length > 0 && <span> → {tourneesPublic.join(', ')}</span>}
          {tourneesPublic.length === 0 && <span> → Vérifiez que <code>public/tournees/liste.json</code> existe</span>}
        </div>

        {/* Tournées disponibles sur le serveur */}
        {tourneesPublic.length > 0 && (
          <div style={{ marginBottom:'20px' }}>
            <h3 style={{ fontSize:'18px', fontWeight:'bold', color:'#374151', marginBottom:'12px' }}>📂 Tournées disponibles</h3>
            <div style={{ display:'grid', gap:'8px' }}>
              {tourneesPublic.map(fichier => (
                <button
                  key={fichier}
                  onClick={() => chargerDepuisPublic(fichier)}
                  style={{ padding:'14px', background:'#f0fdf4', border:'2px solid #86efac', borderRadius:'10px', cursor:'pointer', textAlign:'left', fontSize:'15px', fontWeight:'600', color:'#15803d', display:'flex', alignItems:'center', gap:'10px' }}
                >
                  📄 {fichier}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Import fichier local */}
        <label htmlFor="file-upload" style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', border:'2px dashed #93c5fd', borderRadius:'12px', padding:'16px 20px', textAlign:'center', background:'#eff6ff', cursor:'pointer', marginTop:'8px' }}>
          <FileText size={24} style={{ color:'#2563eb', flexShrink:0 }} />
          <span style={{ fontSize:'15px', fontWeight:'600', color:'#374151' }}>
            {tourneesPublic.length > 0 ? 'Ou importer un fichier local (.xlsx)' : 'Importer un fichier Excel (.xlsx)'}
          </span>
          <input id="file-upload" type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:'none' }} />
        </label>
        {cfg.cheminExcel && (
          <div style={{ marginTop:'16px', background:'#f0f9ff', border:'2px solid #bae6fd', padding:'14px', borderRadius:'10px' }}>
            <p style={{ fontSize:'13px', color:'#0369a1', margin:0 }}>📁 Chemin configuré : <strong>{cfg.cheminExcel}</strong></p>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Écran de récapitulatif + chargement cabris ──────────────────────────────
function VueRecap({ tournee, onDemarrer }) {
  const [ordre, setOrdre] = useState(tournee.clients.map((_, i) => i));
  const [chargementDemarre, setChargementDemarre] = useState(false);
  const [heureChargement, setHeureChargement] = useState(null);

  const clientsOrdres = ordre.map(i => tournee.clients[i]);
  const totalCabris = tournee.clients.reduce((t, c) => t + c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0), 0);

  const monter = (pos) => {
    if (pos === 0) return;
    setOrdre(prev => { const n = [...prev]; [n[pos-1], n[pos]] = [n[pos], n[pos-1]]; return n; });
  };
  const descendre = (pos) => {
    if (pos === ordre.length - 1) return;
    setOrdre(prev => { const n = [...prev]; [n[pos], n[pos+1]] = [n[pos+1], n[pos]]; return n; });
  };

  const BtnOrdre = ({ onClick, label, disabled }) => (
    <button onClick={onClick} disabled={disabled} style={{
      padding:'6px 10px', border:'none', borderRadius:'8px',
      cursor: disabled ? 'default' : 'pointer',
      background: disabled ? '#e5e7eb' : '#6b7280',
      color: disabled ? '#9ca3af' : 'white',
      fontSize:'16px', fontWeight:'bold', minWidth:'34px'
    }}>{label}</button>
  );

  const demarrerChargement = () => {
    const now = new Date();
    setHeureChargement(now);
    setChargementDemarre(true);
  };

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'32px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'26px', fontWeight:'bold', marginBottom:'24px' }}>
          <CheckCircle size={30} style={{ color:'#16a34a', verticalAlign:'middle', marginRight:'10px' }} />
          Clients à livrer
        </h2>

        {/* Compteurs */}
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'12px', marginBottom:'24px' }}>
          {[
            { label:'Clients',      value: tournee.clients.length,                             color:'#2563eb', bg:'#eff6ff' },
            { label:'Total Cabris', value: totalCabris,                                        color:'#16a34a', bg:'#f0fdf4' },
            { label:'Date',         value: new Date(tournee.date).toLocaleDateString('fr-FR'), color:'#7c3aed', bg:'#faf5ff', small: true },
          ].map(({ label, value, color, bg, small }) => (
            <div key={label} style={{ background:bg, borderRadius:'12px', padding:'16px', textAlign:'center' }}>
              <p style={{ color:'#6b7280', marginBottom:'6px', fontSize:'13px' }}>{label}</p>
              <p style={{ fontSize: small ? '15px' : '24px', fontWeight:'bold', color, margin:0, wordBreak:'break-word', lineHeight:1.2 }}>{value}</p>
            </div>
          ))}
        </div>

        {/* Ordre de livraison */}
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'12px' }}>
          <p style={{ fontWeight:'bold', fontSize:'17px', margin:0 }}>🗂️ Ordre de livraison</p>
          <p style={{ fontSize:'12px', color:'#6b7280', margin:0 }}>↑↓ pour réorganiser</p>
        </div>
        <div style={{ marginBottom:'24px' }}>
          {clientsOrdres.map((c, pos) => (
            <div key={c.id} style={{
              background: pos === 0 ? '#f0fdf4' : '#f9fafb',
              border: pos === 0 ? '2px solid #86efac' : '2px solid #e5e7eb',
              borderRadius:'12px', padding:'12px 14px', marginBottom:'8px',
              display:'flex', alignItems:'center', gap:'12px'
            }}>
              <div style={{
                minWidth:'36px', height:'36px', borderRadius:'50%', display:'flex', alignItems:'center', justifyContent:'center',
                background: pos === 0 ? '#16a34a' : '#e5e7eb',
                color: pos === 0 ? 'white' : '#374151', fontWeight:'bold', fontSize:'16px', flexShrink:0
              }}>{pos + 1}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <p style={{ fontWeight:'bold', fontSize:'15px', margin:'0 0 2px', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                  {pos === 0 ? '🚀 ' : ''}{c.nom}
                </p>
                <p style={{ color:'#6b7280', fontSize:'12px', margin:0, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                  {c.adresse} · {c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0)} cabris
                </p>
              </div>
              <div style={{ display:'flex', gap:'6px', flexShrink:0 }}>
                <BtnOrdre onClick={() => monter(pos)}    label="↑" disabled={pos === 0} />
                <BtnOrdre onClick={() => descendre(pos)} label="↓" disabled={pos === ordre.length - 1} />
              </div>
            </div>
          ))}
        </div>

        {/* Phase chargement */}
        {!chargementDemarre ? (
          <button onClick={demarrerChargement}
            style={{ width:'100%', padding:'24px', background:'#f59e0b', color:'white', border:'none', borderRadius:'14px', fontSize:'22px', fontWeight:'bold', cursor:'pointer' }}>
            📦 Démarrer le chargement
          </button>
        ) : (
          <div>
            <div style={{ background:'#fefce8', border:'2px solid #fde68a', borderRadius:'12px', padding:'16px', marginBottom:'16px', textAlign:'center' }}>
              <p style={{ fontWeight:'bold', color:'#92400e', fontSize:'16px', margin:'0 0 4px' }}>⏱️ Chargement en cours...</p>
              <p style={{ color:'#b45309', fontSize:'14px', margin:0 }}>
                Début : {heureChargement.toLocaleTimeString('fr-FR')}
              </p>
            </div>
            <button onClick={() => onDemarrer(clientsOrdres, heureChargement.toISOString())}
              style={{ width:'100%', padding:'24px', background:'#16a34a', color:'white', border:'none', borderRadius:'14px', fontSize:'22px', fontWeight:'bold', cursor:'pointer' }}>
              <Clock size={26} style={{ verticalAlign:'middle', marginRight:'10px' }} />
              DÉPART DE L'UNITÉ
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Écran de navigation vers le client ─────────────────────────────────────
function VueTournee({ client, index, total, onArrivee }) {
  const openWaze = () => {
    window.open(`https://waze.com/ul?q=${encodeURIComponent(client.adresse)}&navigate=yes`, '_blank');
  };
  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#1d4ed8,#1e40af)', padding:'24px', display:'flex', alignItems:'center', justifyContent:'center' }}>
      <div style={{ maxWidth:'560px', width:'100%', background:'white', borderRadius:'20px', padding:'36px', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign:'center', marginBottom:'24px' }}>
          <div style={{ display:'inline-block', background:'#dbeafe', padding:'10px 24px', borderRadius:'999px', marginBottom:'16px' }}>
            <span style={{ color:'#1d4ed8', fontWeight:'bold', fontSize:'17px' }}>Client {index + 1} / {total}</span>
          </div>
          <h2 style={{ fontSize:'28px', fontWeight:'bold', color:'#1f2937', margin:'0 0 8px' }}>{client.nom}</h2>
          <p style={{ color:'#6b7280', fontSize:'17px', margin:0 }}>{client.adresse}</p>
        </div>
        <div style={{ background:'#f9fafb', borderRadius:'12px', padding:'16px', marginBottom:'16px' }}>
          <p style={{ color:'#6b7280', fontSize:'14px', marginBottom:'6px' }}>Services à livrer</p>
          <p style={{ fontSize:'22px', fontWeight:'bold', color:'#2563eb', margin:0 }}>
            {client.services.length} service(s) — {client.services.reduce((s, sv) => s + sv.cabrisPrevu, 0)} cabris
          </p>
        </div>
        {client.commentaireFixe && (
          <div style={{ background:'#fefce8', borderLeft:'4px solid #f59e0b', padding:'14px', borderRadius:'8px', marginBottom:'16px' }}>
            <p style={{ fontWeight:'600', color:'#92400e', marginBottom:'4px', fontSize:'14px' }}>Note importante</p>
            <p style={{ color:'#374151', margin:0 }}>{client.commentaireFixe}</p>
          </div>
        )}
        <div style={{ display:'flex', flexDirection:'column', gap:'12px' }}>
          <button onClick={openWaze} style={{ width:'100%', padding:'22px', background:'#2563eb', color:'white', border:'none', borderRadius:'14px', fontSize:'19px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
            <MapPin size={26} /> Ouvrir Waze
          </button>
          <button onClick={onArrivee} style={{ width:'100%', padding:'22px', background:'#16a34a', color:'white', border:'none', borderRadius:'14px', fontSize:'19px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
            <ArrowRight size={26} /> ARRIVÉE CLIENT
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Écran d'intervention chez le client ────────────────────────────────────
function VueClient({ client, onDepart }) {
  const [services, setServices]           = useState(() => client.services.map(s => ({ ...s, cabrisRecuperes: 0 })));
  const [commentaire, setCommentaire]     = useState('');
  const [photos, setPhotos]               = useState([]);
  const [photoAgrandie, setPhotoAgrandie] = useState(null);
  const refCamera  = React.useRef(null);
  const refGalerie = React.useRef(null);

  const updateCabris = (index, valeur) => {
    setServices(prev => { const n = [...prev]; n[index] = { ...n[index], cabrisRecuperes: parseInt(valeur) || 0 }; return n; });
  };

  const traiterFichiers = (e) => {
    Array.from(e.target.files || []).forEach(file => {
      const reader = new FileReader();
      reader.onload = async (ev) => {
        const compressed = await compresserPhoto(ev.target.result, 800);
        setPhotos(prev => [...prev, { id: `p-${Date.now()}-${Math.random().toString(36).slice(2)}`, data: compressed, heure: new Date().toLocaleTimeString('fr-FR') }]);
      };
      reader.readAsDataURL(file);
    });
    e.target.value = '';
  };

  const supprimerPhoto = (id) => setPhotos(prev => prev.filter(p => p.id !== id));
  const btnStyle = (bg) => ({ width:'100%', padding:'20px', background:bg, color:'white', border:'none', borderRadius:'12px', fontSize:'17px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' });

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'36px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'24px', fontWeight:'bold', color:'#1f2937', marginBottom:'24px' }}>{client.nom}</h2>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'16px', marginBottom:'24px' }}>
          <div style={{ background:'#eff6ff', borderRadius:'12px', padding:'16px' }}>
            <p style={{ fontWeight:'600', color:'#1d4ed8', fontSize:'14px', marginBottom:'10px' }}>Codes d'accès</p>
            {client.code1 && <p style={{ fontFamily:'monospace', fontSize:'18px', margin:'4px 0' }}>🔑 {client.code1}</p>}
            {client.code2 && <p style={{ fontFamily:'monospace', fontSize:'18px', margin:'4px 0' }}>🔑 {client.code2}</p>}
            {client.code3 && <p style={{ fontFamily:'monospace', fontSize:'18px', margin:'4px 0' }}>🔑 {client.code3}</p>}
            {!client.code1 && !client.code2 && !client.code3 && <p style={{ color:'#9ca3af', margin:0 }}>Aucun code</p>}
          </div>
          <div style={{ background:'#f0fdf4', borderRadius:'12px', padding:'16px' }}>
            <p style={{ fontWeight:'600', color:'#15803d', fontSize:'14px', marginBottom:'10px' }}>Contact</p>
            <p style={{ fontSize:'20px', fontWeight:'bold', margin:0 }}>📞 {client.telephone || 'Non renseigné'}</p>
          </div>
        </div>

        <h3 style={{ fontSize:'20px', fontWeight:'bold', marginBottom:'16px' }}>Services à livrer</h3>
        {services.map((sv, i) => (
          <div key={sv.id} style={{ background:'#f9fafb', borderRadius:'12px', padding:'20px', marginBottom:'12px' }}>
            <p style={{ fontWeight:'bold', fontSize:'18px', marginBottom:'14px' }}>{sv.nom}</p>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'16px' }}>
              <div>
                <p style={{ fontSize:'13px', color:'#6b7280', marginBottom:'8px' }}>Cabris livrés (propre)</p>
                <div style={{ background:'#dbeafe', border:'2px solid #93c5fd', borderRadius:'10px', padding:'16px', textAlign:'center' }}>
                  <span style={{ fontSize:'36px', fontWeight:'bold', color:'#1d4ed8' }}>{sv.cabrisPrevu}</span>
                </div>
              </div>
              <div>
                <p style={{ fontSize:'13px', color:'#6b7280', marginBottom:'8px' }}>Cabris récupérés (sale)</p>
                <input type="number" min="0"
                  value={sv.cabrisRecuperes === 0 ? '' : sv.cabrisRecuperes}
                  onChange={e => updateCabris(i, e.target.value)}
                  style={{ width:'100%', padding:'16px', fontSize:'36px', fontWeight:'bold', textAlign:'center', border:'2px solid #d1d5db', borderRadius:'10px', outline:'none', boxSizing:'border-box' }}
                />
              </div>
            </div>
          </div>
        ))}

        <div style={{ marginBottom:'24px' }}>
          <p style={{ fontSize:'18px', fontWeight:'bold', marginBottom:'10px' }}>Commentaire de livraison</p>
          <textarea value={commentaire} onChange={e => setCommentaire(e.target.value)} rows={4}
            placeholder="Incident, observation, demande particulière..."
            style={{ width:'100%', padding:'16px', fontSize:'16px', border:'2px solid #d1d5db', borderRadius:'12px', outline:'none', resize:'vertical', boxSizing:'border-box' }}
          />
        </div>

        <div style={{ marginBottom:'28px', background:'#f0f9ff', border:'2px solid #bae6fd', borderRadius:'16px', padding:'20px' }}>
          <p style={{ fontSize:'18px', fontWeight:'bold', marginBottom:'16px', color:'#0369a1' }}>
            📷 Photos de livraison {photos.length > 0 && `(${photos.length})`}
          </p>
          <input ref={refCamera}  type="file" accept="image/*" capture="environment" onChange={traiterFichiers} style={{ position:'absolute', width:'1px', height:'1px', opacity:0, pointerEvents:'none' }} />
          <input ref={refGalerie} type="file" accept="image/*" multiple onChange={traiterFichiers} style={{ position:'absolute', width:'1px', height:'1px', opacity:0, pointerEvents:'none' }} />
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px', marginBottom:'16px' }}>
            <button onClick={() => refCamera.current.click()}  style={btnStyle('#0284c7')}>📷 Appareil photo</button>
            <button onClick={() => refGalerie.current.click()} style={btnStyle('#7c3aed')}>🖼️ Galerie</button>
          </div>
          {photos.length === 0 ? (
            <div style={{ textAlign:'center', padding:'20px', color:'#94a3b8', border:'2px dashed #cbd5e1', borderRadius:'10px' }}>Aucune photo prise</div>
          ) : (
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'10px' }}>
              {photos.map(photo => (
                <div key={photo.id} style={{ position:'relative', borderRadius:'10px', overflow:'hidden', border:'2px solid #e2e8f0' }}>
                  <img src={photo.data} alt="livraison" onClick={() => setPhotoAgrandie(photo)}
                    style={{ width:'100%', height:'110px', objectFit:'cover', display:'block', cursor:'pointer' }} />
                  <div style={{ position:'absolute', top:0, left:0, right:0, background:'rgba(0,0,0,0.45)', padding:'4px 8px', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                    <span style={{ color:'white', fontSize:'11px' }}>🕐 {photo.heure}</span>
                    <button onClick={(e) => { e.stopPropagation(); supprimerPhoto(photo.id); }}
                      style={{ background:'#ef4444', color:'white', border:'none', borderRadius:'6px', width:'24px', height:'24px', fontSize:'14px', cursor:'pointer', fontWeight:'bold', padding:0 }}>✕</button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <button onClick={() => onDepart(commentaire, services, photos)}
          style={{ width:'100%', padding:'24px', background:'#16a34a', color:'white', border:'none', borderRadius:'14px', fontSize:'22px', fontWeight:'bold', cursor:'pointer' }}>
          <ArrowRight size={26} style={{ verticalAlign:'middle', marginRight:'10px' }} />
          VALIDER ET DÉPART
        </button>
      </div>

      {photoAgrandie && (
        <div onClick={() => setPhotoAgrandie(null)}
          style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.92)', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', zIndex:9999, padding:'16px' }}>
          <img src={photoAgrandie.data} alt="agrandie" style={{ maxWidth:'100%', maxHeight:'80vh', objectFit:'contain', borderRadius:'8px' }} />
          <p style={{ color:'#94a3b8', marginTop:'16px', fontSize:'14px' }}>Prise à {photoAgrandie.heure} · Touchez pour fermer</p>
        </div>
      )}
    </div>
  );
}

// ─── Écran fin de tournée (retour unité) ────────────────────────────────────
function VueFin({ onDecharge, adresseUnite }) {
  const ouvrirWazeUnite = () => {
    const dest = adresseUnite || 'unité de départ';
    window.open(`https://waze.com/ul?q=${encodeURIComponent(dest)}&navigate=yes`, '_blank');
  };

  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#15803d,#166534)', display:'flex', alignItems:'center', justifyContent:'center', padding:'24px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'40px', maxWidth:'480px', width:'100%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <CheckCircle size={80} style={{ color:'#16a34a', margin:'0 auto 20px' }} />
        <h2 style={{ fontSize:'28px', fontWeight:'bold', color:'#1f2937', marginBottom:'12px' }}>
          Tous les clients sont livrés !
        </h2>
        <p style={{ color:'#6b7280', fontSize:'17px', marginBottom:'32px' }}>
          Retournez à l'unité pour décharger le linge sale.
        </p>

        {adresseUnite && (
          <button onClick={ouvrirWazeUnite}
            style={{ width:'100%', padding:'22px', background:'#2563eb', color:'white', border:'none', borderRadius:'14px', fontSize:'19px', fontWeight:'bold', cursor:'pointer', marginBottom:'12px', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
            <MapPin size={24} /> Retour à l'unité via Waze
          </button>
        )}

        <div style={{ background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'12px', padding:'16px', marginBottom:'24px' }}>
          <p style={{ fontWeight:'700', color:'#92400e', margin:'0 0 6px', fontSize:'15px' }}>⏱️ Horodatage en cours</p>
          <p style={{ color:'#b45309', fontSize:'14px', margin:0 }}>
            Le temps de tournée inclut le déchargement.<br/>
            Cliquez ci-dessous une fois le camion déchargé.
          </p>
        </div>

        <button onClick={onDecharge}
          style={{ width:'100%', padding:'24px', background:'#dc2626', color:'white', border:'none', borderRadius:'14px', fontSize:'20px', fontWeight:'bold', cursor:'pointer' }}>
          ✅ Déchargement terminé — Clôturer la tournée
        </button>
      </div>
    </div>
  );
}

// ─── Écran rapport final ─────────────────────────────────────────────────────
function VueRapport({ tournee, clientData, logs, startTime, agentName, onNouvelle }) {
  const cfg = chargerConfig();
  // config est chargé depuis config.json + localStorage via le state

  const logDepart    = logs.find(l => l.type === 'DEPART_UNITE');
  const logChargement = logs.find(l => l.type === 'DEBUT_CHARGEMENT');
  const logFin       = logs.find(l => l.type === 'FIN_DECHARGE');

  const dureeTotal = (startTime && logFin)
    ? Math.round((new Date(logFin.timestamp) - new Date(startTime)) / 60000) : 0;

  const creerDocPDF = () => {
    const doc = new jsPDF();

    // En-tête
    doc.setFontSize(18);
    doc.text('Rapport de Tournee BDL', 105, 20, { align: 'center' });
    doc.setFontSize(11);
    doc.text(`Chauffeur : ${agentName}`, 20, 32);
    doc.text(`Date : ${new Date(tournee.date).toLocaleDateString('fr-FR')}`, 20, 39);
    if (logChargement) doc.text(`Debut chargement : ${new Date(logChargement.timestamp).toLocaleTimeString('fr-FR')}`, 20, 46);
    if (logDepart)     doc.text(`Depart unite      : ${new Date(logDepart.timestamp).toLocaleTimeString('fr-FR')}`, 20, 53);
    if (logFin)        doc.text(`Fin decharge      : ${new Date(logFin.timestamp).toLocaleTimeString('fr-FR')}`, 20, 60);
    doc.text(`Duree totale      : ${Math.floor(dureeTotal/60)}h ${dureeTotal%60}min`, 20, 67);

    let y = 78;
    tournee.clients.forEach((client, i) => {
      const d = clientData[client.id];
      if (!d) return;
      if (y > 250) { doc.addPage(); y = 20; }

      doc.setFontSize(13); doc.setFont(undefined, 'bold');
      doc.text(`${i+1}. ${client.nom}`, 20, y); y += 7;
      doc.setFontSize(10); doc.setFont(undefined, 'normal');
      doc.text(`Adresse : ${client.adresse}`, 25, y); y += 5;
      const arr = d.heureArrivee ? new Date(d.heureArrivee).toLocaleTimeString('fr-FR') : '--';
      const dep = d.heureDepart  ? new Date(d.heureDepart).toLocaleTimeString('fr-FR')  : '--';
      doc.text(`Arrivee : ${arr}  |  Depart : ${dep}`, 25, y); y += 7;

      autoTable(doc, {
        startY: y,
        head: [['Service', 'Livres', 'Recuperes']],
        body: d.services.map(s => [s.nom, String(s.cabrisPrevu), String(s.cabrisRecuperes)]),
        margin: { left: 25 },
        headStyles: { fillColor: [37, 99, 235] },
        styles: { fontSize: 9 }
      });
      y = doc.lastAutoTable.finalY + 5;

      if (d.commentaire) {
        if (y > 270) { doc.addPage(); y = 20; }
        doc.setFont(undefined, 'italic');
        doc.text(`Note : ${d.commentaire}`, 25, y); y += 7;
        doc.setFont(undefined, 'normal');
      }

      if (d.photos && d.photos.length > 0) {
        if (y > 240) { doc.addPage(); y = 20; }
        doc.setFont(undefined, 'bold');
        doc.text(`Photos (${d.photos.length}) :`, 25, y); y += 6;
        doc.setFont(undefined, 'normal');
        d.photos.forEach((photo, pi) => {
          try {
            const col = pi % 2;
            const xPos = col === 0 ? 25 : 100;
            if (col === 0 && pi > 0) y += 88;
            if (y > 240) { doc.addPage(); y = 20; }
            doc.addImage(photo.data, 'JPEG', xPos, y, 60, 80);
            doc.setFontSize(8);
            doc.text(`📷 ${photo.heure}`, xPos, y + 83);
            doc.setFontSize(10);
          } catch {}
        });
        const nbLignes = Math.ceil(d.photos.length / 2);
        y += nbLignes * 88 + 8;
      }
      y += 8;
    });

    return doc;
  };

  const genererPDF = () => {
    const doc = creerDocPDF();
    doc.save(`Rapport_${new Date().toISOString().split('T')[0]}.pdf`);
  };

  // Charger EmailJS dynamiquement
  const chargerEmailJS = () => new Promise((resolve, reject) => {
    if (window.emailjs) { resolve(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js';
    s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });

  // Générer le PDF et envoyer par email avec pièce jointe
  const genererEtEnvoyer = async () => {
    const emails = [cfg.email1, cfg.email2, cfg.email3].filter(Boolean);
    if (emails.length === 0) {
      alert('Aucun email configuré. Allez dans la Configuration.');
      return;
    }

    const nomFichier = `Rapport_${new Date().toISOString().split('T')[0]}.pdf`;
    const doc = creerDocPDF();

    // Si EmailJS configuré → envoi direct avec pièce jointe
    if (cfg.emailjsServiceId && cfg.emailjsTemplateId && cfg.emailjsPublicKey) {
      try {
        await chargerEmailJS();
        window.emailjs.init(cfg.emailjsPublicKey);
        const pdfBase64 = doc.output('datauristring').split(',')[1];
        let envoyes = 0;
        for (const email of emails) {
          await window.emailjs.send(cfg.emailjsServiceId, cfg.emailjsTemplateId, {
            to_email: email,
            from_name: 'BDL-LIVRAISON',
            chauffeur: agentName,
            date: new Date(tournee.date).toLocaleDateString('fr-FR'),
            duree: `${Math.floor(dureeTotal/60)}h ${dureeTotal%60}min`,
            nb_clients: String(tournee.clients.length),
            pdf_base64: pdfBase64,
            pdf_name: nomFichier,
          });
          envoyes++;
        }
        doc.save(nomFichier);
        alert(`✅ Rapport envoyé à ${envoyes} destinataire(s) et téléchargé !`);
      } catch (err) {
        console.error('Erreur EmailJS:', err);
        alert('Erreur envoi EmailJS. Vérifiez la configuration.\nLe PDF va être téléchargé localement.');
        doc.save(nomFichier);
      }
    } else {
      // Fallback : télécharger PDF + ouvrir mailto
      doc.save(nomFichier);
      setTimeout(() => {
        const sujet = encodeURIComponent(`Rapport tournée BDL - ${agentName} - ${new Date(tournee.date).toLocaleDateString('fr-FR')}`);
        const corps = encodeURIComponent(
          `Bonjour,\n\nVeuillez trouver en pièce jointe le rapport de tournée (${nomFichier}).\n\n` +
          `Chauffeur : ${agentName}\n` +
          `Date : ${new Date(tournee.date).toLocaleDateString('fr-FR')}\n` +
          `Durée : ${Math.floor(dureeTotal/60)}h ${dureeTotal%60}min\n` +
          `Clients : ${tournee.clients.length}\n\nCordialement`
        );
        window.open(`mailto:${emails.join(',')}?subject=${sujet}&body=${corps}`);
      }, 800);
    }
  };

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'36px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'28px', fontWeight:'bold', textAlign:'center', marginBottom:'24px' }}>
          <FileText size={32} style={{ verticalAlign:'middle', marginRight:'10px' }} />
          Rapport de Tournée
        </h2>

        {/* Horaires récap */}
        <div style={{ background:'#f9fafb', border:'2px solid #e5e7eb', borderRadius:'12px', padding:'16px', marginBottom:'24px' }}>
          <p style={{ fontWeight:'bold', fontSize:'15px', marginBottom:'10px', color:'#374151' }}>⏱️ Chronologie</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px', fontSize:'14px', color:'#6b7280' }}>
            {logChargement && <span>📦 Début chargement : <strong>{new Date(logChargement.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {logDepart     && <span>🚚 Départ unité : <strong>{new Date(logDepart.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {logFin        && <span>✅ Fin décharge : <strong>{new Date(logFin.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            <span>⏳ Durée totale : <strong>{Math.floor(dureeTotal/60)}h {dureeTotal%60}min</strong></span>
          </div>
        </div>

        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'12px', marginBottom:'24px' }}>
          {[
            { label:'Clients', value: tournee.clients.length, color:'#16a34a', bg:'#f0fdf4' },
            { label:'Chauffeur', value: agentName, color:'#7c3aed', bg:'#faf5ff' },
            { label:'Durée', value:`${Math.floor(dureeTotal/60)}h ${dureeTotal%60}min`, color:'#2563eb', bg:'#eff6ff' },
          ].map(({ label, value, color, bg }) => (
            <div key={label} style={{ background:bg, borderRadius:'12px', padding:'14px', textAlign:'center' }}>
              <p style={{ color:'#6b7280', fontSize:'12px', marginBottom:'4px' }}>{label}</p>
              <p style={{ fontSize:'18px', fontWeight:'bold', color, margin:0, wordBreak:'break-word' }}>{value}</p>
            </div>
          ))}
        </div>

        <button onClick={genererEtEnvoyer}
          style={{ width:'100%', padding:'22px', background:'linear-gradient(135deg,#2563eb,#0891b2)', color:'white', border:'none', borderRadius:'14px', fontSize:'20px', fontWeight:'bold', cursor:'pointer', marginBottom:'12px', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', boxShadow:'0 4px 14px rgba(37,99,235,0.4)' }}>
          <Download size={24} /> 📧 Générer PDF et envoyer par email
        </button>
        <button onClick={genererPDF}
          style={{ width:'100%', padding:'14px', background:'#f3f4f6', color:'#374151', border:'2px solid #e5e7eb', borderRadius:'12px', fontSize:'15px', fontWeight:'600', cursor:'pointer', marginBottom:'12px', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px' }}>
          <Download size={18} /> Télécharger le PDF uniquement
        </button>

        <button onClick={onNouvelle}
          style={{ width:'100%', padding:'22px', background:'#4b5563', color:'white', border:'none', borderRadius:'14px', fontSize:'20px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
          <Home size={24} /> Nouvelle tournée
        </button>

        {/* Détail interventions */}
        <div style={{ borderTop:'2px solid #e5e7eb', marginTop:'24px', paddingTop:'24px' }}>
          <h3 style={{ fontSize:'20px', fontWeight:'bold', marginBottom:'16px' }}>Détail des interventions</h3>
          {tournee.clients.map((c, i) => {
            const d = clientData[c.id];
            if (!d) return null;
            return (
              <div key={c.id} style={{ background:'#f9fafb', borderRadius:'12px', padding:'16px', marginBottom:'12px' }}>
                <div style={{ display:'flex', justifyContent:'space-between', marginBottom:'10px' }}>
                  <p style={{ fontWeight:'bold', fontSize:'17px', margin:0 }}>{i+1}. {c.nom}</p>
                  <div style={{ textAlign:'right', fontSize:'13px', color:'#6b7280' }}>
                    <div>{d.heureArrivee ? new Date(d.heureArrivee).toLocaleTimeString('fr-FR') : '--'}</div>
                    <div>→ {d.heureDepart ? new Date(d.heureDepart).toLocaleTimeString('fr-FR') : '--'}</div>
                  </div>
                </div>
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px', marginBottom:'8px' }}>
                  {d.services.map(s => (
                    <div key={s.id} style={{ background:'white', padding:'10px', borderRadius:'8px' }}>
                      <p style={{ fontWeight:'600', margin:'0 0 4px', fontSize:'14px' }}>{s.nom}</p>
                      <p style={{ color:'#6b7280', fontSize:'13px', margin:0 }}>Livrés: {s.cabrisPrevu} | Récupérés: {s.cabrisRecuperes}</p>
                    </div>
                  ))}
                </div>
                {d.commentaire && <p style={{ fontSize:'13px', fontStyle:'italic', color:'#6b7280', background:'#fefce8', padding:'8px', borderRadius:'6px', margin:'0 0 8px' }}>💬 {d.commentaire}</p>}
                {d.photos && d.photos.length > 0 && (
                  <div>
                    <p style={{ fontSize:'13px', fontWeight:'600', color:'#374151', margin:'0 0 6px' }}>📷 {d.photos.length} photo(s)</p>
                    <div style={{ display:'flex', gap:'6px', flexWrap:'wrap' }}>
                      {d.photos.map(photo => (
                        <img key={photo.id} src={photo.data} alt="livraison" style={{ width:'64px', height:'64px', objectFit:'cover', borderRadius:'6px', border:'2px solid #e5e7eb' }} />
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─── Composant principal ─────────────────────────────────────────────────────
export default function App() {
  const [vue, setVue]                   = useState('login');
  const [agentName, setAgentName]       = useState('');
  const [tournee, setTournee]           = useState(null);
  const [clientIndex, setClientIndex]   = useState(0);
  const [startTime, setStartTime]       = useState(null);
  const [logs, setLogs]                 = useState([]);
  const [clientData, setClientData]     = useState({});
  const [config, setConfig]             = useState(CONFIG_DEFAULT);

  // Chargement de la config au démarrage (depuis config.json + localStorage)
  useEffect(() => {
    chargerConfigAsync().then(cfg => setConfig(cfg));
  }, []);

  // Restauration localStorage
  useEffect(() => {
    try {
      const raw = localStorage.getItem('bdl');
      if (raw) {
        const s = JSON.parse(raw);
        if (s.vue && s.vue !== 'login') {
          setVue(s.vue); setAgentName(s.agentName || '');
          setTournee(s.tournee || null); setClientIndex(s.clientIndex || 0);
          setStartTime(s.startTime || null); setLogs(s.logs || []);
          setClientData(s.clientData || {});
        }
      }
    } catch {}
  }, []);

  // Sauvegarde auto (sans photos)
  useEffect(() => {
    if (vue !== 'login') {
      try {
        const clientDataSansPhotos = Object.fromEntries(
          Object.entries(clientData).map(([k, v]) => [k, { ...v, photos: [] }])
        );
        localStorage.setItem('bdl', JSON.stringify({ vue, agentName, tournee, clientIndex, startTime, logs, clientData: clientDataSansPhotos }));
      } catch {
        try { localStorage.setItem('bdl', JSON.stringify({ vue, agentName, clientIndex, startTime, logs, tournee, clientData: {} })); } catch {}
      }
    }
  }, [vue, agentName, tournee, clientIndex, startTime, logs, clientData]);

  const addLog = (type, idx = null, extra = {}) => {
    const entry = { type, clientIndex: idx, timestamp: new Date().toISOString(), ...extra };
    setLogs(prev => [...prev, entry]);
    return entry;
  };

  const handleLogin    = (nom) => { setAgentName(nom); setVue('import'); };
  const handleConfig   = ()    => setVue('config');
  const handleImport   = (t)   => { setTournee(t); setClientIndex(0); setClientData({}); setLogs([]); setVue('recap'); };

  const handleDemarrer = (clientsOrdres, heureChargement) => {
    setTournee(prev => ({ ...prev, clients: clientsOrdres }));
    const now = new Date().toISOString();
    setStartTime(heureChargement); // l'horodatage démarre dès le chargement
    setLogs([
      { type: 'DEBUT_CHARGEMENT', clientIndex: null, timestamp: heureChargement },
      { type: 'DEPART_UNITE',     clientIndex: null, timestamp: now }
    ]);
    setVue('tournee');
  };

  const handleArrivee = () => { addLog('ARRIVEE_CLIENT', clientIndex); setVue('client'); };

  const handleDepartClient = (commentaire, services, photos) => {
    const now = new Date().toISOString();
    const client = tournee.clients[clientIndex];
    const heureArrivee = [...logs].reverse().find(l => l.type === 'ARRIVEE_CLIENT' && l.clientIndex === clientIndex)?.timestamp;
    setClientData(prev => ({ ...prev, [client.id]: { services, commentaire, photos: photos || [], heureArrivee, heureDepart: now } }));
    addLog('DEPART_CLIENT', clientIndex, { commentaire });
    if (clientIndex < tournee.clients.length - 1) {
      setClientIndex(clientIndex + 1); setVue('tournee');
    } else {
      setVue('fin');
    }
  };

  const handleDecharge = () => { addLog('FIN_DECHARGE'); setVue('rapport'); };
  const handleNouvelle = () => { localStorage.removeItem('bdl'); setVue('login'); setTournee(null); setAgentName(''); setLogs([]); setClientData({}); };

  // config est chargé depuis config.json + localStorage via le state

  if (vue === 'config')  return <VueConfig config={config} onSave={(cfg) => { setConfig(cfg); }} onRetour={() => setVue(agentName ? 'import' : 'login')} />;
  if (vue === 'login')   return <VueLogin  onLogin={handleLogin} onConfig={handleConfig} />;
  if (vue === 'import')  return <VueImport key={Date.now()} onImport={handleImport} onConfig={handleConfig} />;
  if (vue === 'recap')   return <VueRecap  tournee={tournee} onDemarrer={handleDemarrer} />;
  if (vue === 'fin')     return <VueFin    onDecharge={handleDecharge} adresseUnite={config.adresseUnite} />;
  if (vue === 'rapport') return <VueRapport tournee={tournee} clientData={clientData} logs={logs} startTime={startTime} agentName={agentName} onNouvelle={handleNouvelle} />;

  if (vue === 'tournee' && tournee)
    return <VueTournee client={tournee.clients[clientIndex]} index={clientIndex} total={tournee.clients.length} onArrivee={handleArrivee} />;
  if (vue === 'client' && tournee)
    return <VueClient key={`client-${clientIndex}`} client={tournee.clients[clientIndex]} onDepart={handleDepartClient} />;

  return null;
}
