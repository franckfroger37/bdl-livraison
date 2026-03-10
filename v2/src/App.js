import React, { useState, useEffect } from 'react';
import { MapPin, Clock, Package, FileText, Download, ArrowRight, Home, CheckCircle, Settings } from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// ─── Configuration ───────────────────────────────────────────────────────────
const APP_VERSION = '2.0.1';

// ─── Helper : parser les lignes Excel (format ID_Cabri — une ligne = un cabri) ─
// Règles :
//   • Colonne Client vide → même client que la ligne précédente (carry-forward)
//   • Colonne Service vide → même service que la ligne précédente (carry-forward)
//   • Chaque ID_Cabri est ajouté au service courant
//   • cabrisPrevu = nombre d'IDs du service (calculé automatiquement)
function parserRowsExcel(rows) {
  const clients = [];
  let current = null;
  let dernierServiceNom = '';

  rows.forEach(row => {
    const clientVal  = row.Client   ? String(row.Client).trim()   : '';
    const serviceVal = row.Service  ? String(row.Service).trim()  : '';
    const idCabriVal = (row.ID_Cabri !== undefined && row.ID_Cabri !== '')
      ? String(row.ID_Cabri).trim() : '';

    // Ignorer les lignes entièrement vides
    if (!clientVal && !serviceVal && !idCabriVal) return;

    // Nouveau client
    if (clientVal) {
      if (current) clients.push(current);
      current = {
        id: `C${clients.length + 1}`,
        nom: clientVal,
        adresse:    row.Adresse   ? String(row.Adresse).trim()   : '',
        gps:        row.GPS       ? String(row.GPS).trim()       : '',
        telephone:  row.Telephone ? String(row.Telephone).trim() : '',
        noteTournee: row['Note de tournée']
          ? String(row['Note de tournée']).trim()
          : (row['Note de tounée'] ? String(row['Note de tounée']).trim() : ''),
        services: []
      };
      dernierServiceNom = '';
    }

    if (!current) return;

    // Note de tournée peut apparaître sur n'importe quelle ligne du client
    const note = row['Note de tournée'] || row['Note de tounée'] || '';
    if (note && !current.noteTournee) current.noteTournee = String(note).trim();

    // Nom du service : nouveau si non vide, sinon carry-forward
    const nomService = serviceVal || dernierServiceNom;
    if (!nomService) return;

    // Créer un nouveau service si le nom change
    if (nomService !== dernierServiceNom) {
      current.services.push({
        id: `S${current.services.length + 1}`,
        nom: nomService,
        cabrisIds: [],
        cabrisPrevu: 0,
        cabrisRecuperes: 0
      });
      dernierServiceNom = nomService;
    }

    // Ajouter l'ID cabri au service courant (si présent)
    if (idCabriVal) {
      const svc = current.services[current.services.length - 1];
      svc.cabrisIds.push(idCabriVal);
      svc.cabrisPrevu = svc.cabrisIds.length; // toujours égal au nombre d'IDs
    }
  });

  if (current) clients.push(current);
  return clients;
}
const LOGO_BDL = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAeAB4AAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8lJCIfIiEmKzcvJik0KSEiMEExNDk7Pj4+JS5ESUM8SDc9Pjv/2wBDAQoLCw4NDhwQEBw7KCIoOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozv/wAARCABvAIEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD2aiiigAoorC8T+JI9Ct1SNVku5RlEPRR/ePt/Omk5OyA23dI1LSMqKOpY4FQx39nK2yO7gdvRZATXn0GgeIPE+Ly7mKxvyrTkgEf7KDoPyqS5+H2pQxl7ee3uGH8IyhP0zxW3s4LRyHY9ForzLSfE2p6Bd/ZbwSywodskEv3k/wB0n+XQ16RbXMN5bR3Nu4eKVQysO4qJ03AGiWiiisxBRRRQAUUUUAFFFFABRRRQAUUUUAFee6VAvibxpc3VyN8EDFtp6EA4QfTvXoVcF4Jf7F4h1CxkwsrBgoPcqx4/I5ranpGTQ0d7RWFH4w0oM8V5I9nPESskUiE4I9CBzWTrHj6FYzFpKF5D/wAtpFwq/Qd/xqo4arJ2UTiqY7DwjzOS/X7ix480mO40z+0kUCa2IDEfxITjn6E5/Oq/w8vme2urB2yImEiewbqPzH61Xe9vP+EBurjU7h5ZL2TbCH64JHT8iaT4dQsbq+uP4Aip+JJP9KuUeWEovodNKp7Smp2tc7yiiiuQsKKKKACiiigAooooAKKKKACiiuA8calfWmtpHbXk8KeQp2xyFRnLc8VtRourPlTObFYhYen7Rq539cP4v0a6stQXX9N3AqQ0uwcow/ix6EcGuu0t2k0mzkkYs7QIWYnJJ2jmrJGRg8ioTcJHRGV0mcA+qeHfEiK+rK9hehcGaP7rfjg/r+dMSz8H6afPl1CTUCvKwqMg/XA/ma6HUfBOk30jSxq9pI3J8kjaf+Anj8sVQj+HVoHBk1Cd19FRVP5811Rr2jZSaXYwlhMPOfPKKuc3quq3vifUYYYICEX5be3Tt7n/ADgV6F4f0ddE0pLXIaUnfKw7sf6dvwp+l6Jp+joVs7cKzDDSMcs31NaFc86ia5Y7HSFFcX49v7yzuLJbW6mgDI5YRuVzyPSt/wAMTS3Hh2zlmkaSRkO53OSeT3qpUXGkql9zjhiozxEqNtV/wP8AM1aK4bx3qF7aalapbXc0CtCSRHIVBOfaup0CWSbQbGWV2kdoVLMxyScd6J0XGmql9wp4qNSvKilrE0KKKKwOsKKKKACiiigArzjx/wD8jBH/ANe6/wA2r0evOPiB/wAh9P8Ar2X+bV3YD+N8jys3/wB2+aI4/G2rR28UFtHAiQxqg+QsTgYyea1NH8fNJOkOqRRorHHnx5AX6j0966jRIYotFsxHGqBoEJ2rjJKjJrifHtjb2uqQTQxrGZ4yXCjAJB6/rW8HRrTdPkt5nJVjisNSVZVL7aHeahctaabc3UYVmiiZ1B6HAzXPeF/FV5rmoyW1xBBGqxFwYwc5yB3PvU9pM8/w/Mkhy32J1yfYEf0rnfh7/wAhyb/r2P8A6EtYwox9lUutUdNXEzeIoqLspK9j0C5lMNtLKoBKIWAPsK5bw34uvdZ1VbSe3gRDGzZQNnI+prpr/wD5B9z/ANcm/ka878Cf8jIn/XF/6VNCnGVGcmtUXi61SGJpRi9G9TR+I3/HzYf7j/zFdH4T/wCRYsf9w/8AoRrnPiN/x82H+4/8xXR+E/8AkWLH/cP/AKEaur/ukPX/ADMqH/Ixq+n+Ry3xE/5Ctp/1wP8A6Ea67w5/yLlh/wBcF/lXI/ET/kK2n/XA/wDoRrrvDn/IuWH/AFwX+VFb/dYBhf8Af6v9dhmv+ILbQbZWkUyzSf6uIHGfcnsK42Tx5rUrloo4EUfwiMtj6nNReOpXfxJIpORHEiqPTjP9a6vTdc8Oafp8NtDewIqoAwCnk45J461pGnClSjLk5mzKdepXxE4e05Ix/Eo6F45F5cpa6lEkTSHaksf3c+hB6fWuwrynxS+nS6w02lujRSIGbyxgB+c/0NenWEjTadbSucs8SsT7kCscVSjFRnFWv0OnL8RUnKdKb5uXr3LFFFFcJ6oV5x8QP+Q+n/Xsv82r0euR8W+F7/V79LyzaJtsQQxs208Enjt3rswc4wq3k7Hm5nSnUw/LBXdzotI/5A1j/wBe8f8A6CK4r4hTxvqdtCrAvFES4HbJ4/lVRdI8XWq+TGt6iDgLHP8AL+hq1pXgfULu5E2qnyYs5dd+53/Lp9a6adOnRn7WU0zgrVa+JpKhGk1tq/I6DTYJD4BEWDvezfA+oJH865fwHcRw+ICkjBTNCyJnucg4/Q16QiLHGsaKFRRtCjoBXB634Hu0unuNJ2yRM24RbtrIfYnqKzoVYSU4TduY6MXh6lN0qlNc3JpY7a//AOQfc/8AXJv5GvO/An/IyJ/1xf8ApQdK8XzL5LLesnTa0/H/AKFW34U8KX+l6iL68aJMIVEancefU9KtRhRozi5JtmMp1cViaclTaUX1KvxGB+0WDY42uP1Wuh8IOr+F7LaQdqsD7Hcad4l0Ia7pwiRgk8TbombpnuD7GuJg0HxRYy+VBDcxqWBbypgFPvwamHJWw6p81mi6ntcNjJVVByUl0+X+Re+IgP8AadoexhI/8erq/DLq/huwKnIEIB+o4NQeKPD/APbtknksqXMBJjLdCD1B/SuRsNE8TWV3FGkNzFEJVLiOYBCM8ng4oXJWw6jzWaCXtcNjJVORyUuxH44Ur4mmJHWNCPfit+08C6PdWkVwl1dMsiBgVdccj6VoeJ/DK67GksLiK7iGFZujD0P+NclHonizTsw2yXSIT0hnG0/rWsKntKUYxnytGFWg6OInKpS54y1Vuhuy+CNBhljhlvp0klOI0aVQW+gxXVwQrb28cCZKxoEXPXAGK4HT/Ber3d2tzqM5t8EMXMm+U49PT869BAwAMk4HU9648S9lz8x6WBivel7Lk/UWiiiuQ9EKKKKACsnVvENvpdxHarBNdXUo3LDAuTj1Na1cnd3UOl+PftN84hgmtNqSMPlzkd/wrajBSk7q9kcuJqSpxVna7Sv2G6BqBu/FmqXMkc1unkKTHNwUxjOR26VZPja05mWwvWsw203Qj+SqFtNHqXiLxA9nIJ1lstqFDncdoHH41l2M0L+HRbXPiT7LEFZZLQ24ZhyeB3Oa7nShJ3ku2mvbyPLWIqU48sX1k76a6+bSt6HosUqTRJLGwZHUMrDuD0rnPBjs9nqO9mbF445PTgVt6VClvpNpDHL5qJCoV8Y3DHBx2rnfCF3Bby6jp00gS6a7kZYm6kY7fka5IL3JpeR6FSX72k5aXv8AkS6Rq+m6V4ZkvY/tkluJ2XExDOWP6Yqxb+LoHuoYLywu7ITnbFJMmFY9q5cEf8K5l5/5ff6it/xpj+w7A9/tMfP4GumVKDnZrdtHFDEVVS5ouyjFO1t73NHVPElvp14tlHbT3l0y7jFAuSo96NN8S2l+t15kM1o1ou6ZZ1xtH+RWTcXsOh+N7q71DdFb3Vuqxy7SRkbcjj6Gs+GZdUk8TzWe6VZoVKYBBYc9vwqFQi4bdFr6tGssXUVS11u1y+STs++ptL41tCUlksbyKzdtq3TR/JXRghgCDkHkEV51HLbXXh6K3uvE/lwlFVrUWwZlweBxycHvXoNtGIrWKMNuCIFDY64HWs8RThC3L59/1NsHXqVb8zvou36N/iS0UUVyHoBRRRQAUUUUAFRXFrb3SeXcQRzJ/dkUMP1qWimnYTSasyG3tLa1BW3t4oQeojQLn8qY2m2Lz/aGsrdpc58wxDd+dWaKOZ73FyRtawVB9jtTci6+zxeeBjzdg3fn1qeihNobSe5B9itPJMH2WHyicmPyxtz64p8sEM6BJoUkVTkB1BANSUUXYuVdjG1eDXXuVfTns5LfbgwXCHr65pvh/RrnT5Lu8vpY3u7tgXEQ+VQOgH51t0Vp7V8nKZewj7T2jbv66FYabYrP562VuJc53iIbs/WrNFFZtt7myilsgooopDCiiigD/9k=';

const LOGO_BDL_PDF = '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';

// Dossiers Google Drive par unité
const DOSSIERS_DRIVE = {
  'BDL12': '1SDu8PBRpHP7oTiT3U1CXXH82ScZYou7i',
  'BDL01': '1-_66pBDrVUy3XTrkrybzzpx6y5JTpPWp',
};

const CONFIG_DEFAULT = {
  nomUnite: 'BDL12',
  dossierDriveId: '1SDu8PBRpHP7oTiT3U1CXXH82ScZYou7i',
  email1: '',
  email2: '',
  email3: '',
  adresseUnite: '',
  gpsUnite: '',
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
  const [cfg, setCfg]           = useState({ ...CONFIG_DEFAULT, ...config });
  const [mdp, setMdp]           = useState('');
  const [debloque, setDebloque] = useState(false);
  const [errMdp, setErrMdp]     = useState(false);

  // Synchroniser avec config.json à l'ouverture
  useEffect(() => {
    chargerConfigAsync().then(cfgFichier => {
      setCfg(prev => ({ ...prev, ...cfgFichier }));
    });
  }, []);

  const verifierMdp = () => {
    if (mdp === '6305') { setDebloque(true); setErrMdp(false); }
    else { setErrMdp(true); setMdp(''); }
  };

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

  if (!debloque) return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#1d4ed8,#1e40af)', display:'flex', alignItems:'center', justifyContent:'center', padding:'24px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'40px', width:'100%', maxWidth:'380px', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <img src={LOGO_BDL} alt="BDL" style={{ height:'80px', margin:'0 auto 16px', display:'block' }} />
        <h2 style={{ fontSize:'22px', fontWeight:'bold', color:'#1f2937', marginBottom:'8px' }}>Configuration</h2>
        <p style={{ color:'#6b7280', fontSize:'14px', marginBottom:'24px' }}>Accès administrateur requis</p>
        <input type="password" value={mdp} onChange={e => setMdp(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && verifierMdp()}
          placeholder="Mot de passe"
          style={{ width:'100%', padding:'16px', fontSize:'20px', textAlign:'center', letterSpacing:'8px', border: errMdp ? '2px solid #ef4444' : '2px solid #d1d5db', borderRadius:'12px', outline:'none', boxSizing:'border-box', marginBottom:'8px' }} />
        {errMdp && <p style={{ color:'#ef4444', fontSize:'14px', margin:'0 0 12px' }}>Mot de passe incorrect</p>}
        <button onClick={verifierMdp}
          style={{ width:'100%', padding:'16px', background:'#2563eb', color:'white', border:'none', borderRadius:'12px', fontSize:'17px', fontWeight:'bold', cursor:'pointer', marginBottom:'12px' }}>
          🔓 Déverrouiller
        </button>
        <button onClick={onRetour}
          style={{ width:'100%', padding:'12px', background:'#f3f4f6', color:'#6b7280', border:'none', borderRadius:'12px', fontSize:'15px', cursor:'pointer' }}>
          Retour
        </button>
      </div>
    </div>
  );

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'600px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'32px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <div style={{ textAlign:'center', marginBottom:'20px' }}>
          <img src={LOGO_BDL} alt="BDL" style={{ height:'60px', margin:'0 auto 12px', display:'block' }} />
          <h2 style={{ fontSize:'22px', fontWeight:'bold', margin:0 }}>
            <Settings size={22} style={{ verticalAlign:'middle', marginRight:'8px', color:'#6b7280' }} />
            Configuration
          </h2>
        </div>

        {/* Nom de l'unité */}
        <div style={{ background:'#eff6ff', border:'2px solid #93c5fd', borderRadius:'12px', padding:'16px', marginBottom:'20px' }}>
          <p style={{ fontWeight:'700', color:'#1d4ed8', marginBottom:'12px', fontSize:'15px' }}>🏭 Nom de l'unité</p>
          <div style={{ marginBottom:'12px' }}>
            <label style={{ display:'block', fontWeight:'600', color:'#374151', marginBottom:'6px', fontSize:'15px' }}>Unité (ex: BDL12)</label>
            <input type="text" value={cfg.nomUnite}
              onChange={e => {
                const val = e.target.value.toUpperCase();
                const driveId = DOSSIERS_DRIVE[val] || cfg.dossierDriveId;
                setCfg(prev => ({ ...prev, nomUnite: val, dossierDriveId: driveId }));
              }}
              placeholder="BDL12"
              style={{ width:'100%', padding:'14px', fontSize:'18px', fontWeight:'bold', border:'2px solid #93c5fd', borderRadius:'10px', outline:'none', boxSizing:'border-box' }} />
          </div>
          {champ('ID dossier Google Drive', 'dossierDriveId', 'Ex: 1SDu8PBRpHP7oTiT3U1CXXH82ScZYou7i')}
          {DOSSIERS_DRIVE[cfg.nomUnite] && (
            <p style={{ fontSize:'12px', color:'#16a34a', margin:0 }}>✅ Dossier Drive reconnu pour {cfg.nomUnite}</p>
          )}
        </div>

        {/* Emails */}
        <div style={{ background:'#fdf4ff', border:'2px solid #e9d5ff', borderRadius:'12px', padding:'16px', marginBottom:'20px' }}>
          <p style={{ fontWeight:'700', color:'#7c3aed', marginBottom:'12px', fontSize:'15px' }}>📧 Envoi du rapport PDF</p>
          {champ('Email 1', 'email1', 'responsable@bdl.fr', 'email')}
          {champ('Email 2 (optionnel)', 'email2', 'chef@bdl.fr', 'email')}
          {champ('Email 3 (optionnel)', 'email3', 'archive@bdl.fr', 'email')}
        </div>

        {/* Unité GPS */}
        <div style={{ background:'#f0fdf4', border:'2px solid #86efac', borderRadius:'12px', padding:'16px', marginBottom:'28px' }}>
          <p style={{ fontWeight:'700', color:'#15803d', marginBottom:'12px', fontSize:'15px' }}>🏢 Adresse de l'unité</p>
          {champ("Adresse", 'adresseUnite', 'Ex: 12 Rue de la Blanchisserie, 37000 Tours')}
          {champ('GPS (optionnel)', 'gpsUnite', 'Ex: 47.3941, 0.6848')}
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
  const [nom, setNom]         = useState('');
  const [apropos, setApropos] = useState(false);
  const cfg = chargerConfig();

  if (apropos) return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#1d4ed8,#1e40af)', display:'flex', alignItems:'center', justifyContent:'center', padding:'16px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'36px', width:'100%', maxWidth:'420px', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign:'center', marginBottom:'24px' }}>
          <img src={LOGO_BDL} alt="Bulle de Linge" style={{ height:'100px', margin:'0 auto 16px', display:'block' }} />
          <h1 style={{ fontSize:'26px', fontWeight:'bold', color:'#1f2937', margin:'0 0 4px' }}>BDL-LIVRAISON</h1>
          <p style={{ color:'#6b7280', margin:0 }}>Version {APP_VERSION} — Février 2026</p>
        </div>
        <div style={{ background:'#f9fafb', borderRadius:'12px', padding:'16px', marginBottom:'16px', fontSize:'14px', color:'#374151', lineHeight:'1.8' }}>
          <p style={{ fontWeight:'bold', marginBottom:'8px', margin:'0 0 8px' }}>📋 Nouveautés v{APP_VERSION}</p>
          <p style={{ margin:'0 0 4px' }}>• Filtrage tournées par unité Drive</p>
          <p style={{ margin:'0 0 4px' }}>• Configuration protégée par mot de passe</p>
          <p style={{ margin:'0 0 4px' }}>• Fenêtre À propos avec logo</p>
          <p style={{ margin:'0 0 4px' }}>• Cabris repris unifiés par client</p>
          <p style={{ margin:'0 0 4px' }}>• Popup plein avant clôture tournée</p>
          <p style={{ margin:0 }}>• Heure dans le nom du fichier PDF</p>
        </div>
        <p style={{ textAlign:'center', color:'#9ca3af', fontSize:'13px', marginBottom:'20px' }}>
          © 2026 Bulle de Linge — Tous droits réservés
        </p>
        <button onClick={() => setApropos(false)}
          style={{ width:'100%', padding:'16px', background:'#2563eb', color:'white', border:'none', borderRadius:'12px', fontSize:'17px', fontWeight:'bold', cursor:'pointer' }}>
          ← Retour
        </button>
      </div>
    </div>
  );

  // Vérifier si une session est sauvegardée
  const sessionSauvegardee = (() => {
    try {
      const raw = localStorage.getItem('bdl-session');
      if (!raw) return null;
      const s = JSON.parse(raw);
      return s.tournee && s.vue ? s : null;
    } catch { return null; }
  })();

  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#1d4ed8,#1e40af)', display:'flex', alignItems:'center', justifyContent:'center', padding:'16px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'40px', width:'100%', maxWidth:'420px', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign:'center', marginBottom:'24px' }}>
          <img src={LOGO_BDL} alt="Bulle de Linge" style={{ height:'90px', margin:'0 auto 12px', display:'block' }} />
          <h1 style={{ fontSize:'28px', fontWeight:'bold', color:'#1f2937', margin:0 }}>BDL-LIVRAISON</h1>
          <p style={{ color:'#6b7280', marginTop:'4px', fontSize:'14px' }}>{cfg.nomUnite || 'Gestion de tournees'}</p>
          <div style={{ display:'inline-block', background:'#dbeafe', borderRadius:'6px', padding:'3px 10px', marginTop:'6px' }}>
            <p style={{ color:'#1d4ed8', margin:0, fontSize:'13px', fontWeight:'700', letterSpacing:'0.05em' }}>v{APP_VERSION}</p>
          </div>
        </div>

        {/* Bannière reprise de session */}
        {sessionSauvegardee && (
          <div style={{ background:'#fef3c7', border:'2px solid #f59e0b', borderRadius:'12px', padding:'14px', marginBottom:'20px' }}>
            <p style={{ fontWeight:'bold', color:'#92400e', fontSize:'14px', margin:'0 0 4px' }}>
              Tournee en cours detectee !
            </p>
            <p style={{ color:'#78350f', fontSize:'13px', margin:'0 0 10px' }}>
              {sessionSauvegardee.nomTournee || 'Tournee'} — {sessionSauvegardee.agentName}
            </p>
            <button onClick={() => onLogin(sessionSauvegardee.agentName, true)}
              style={{ width:'100%', padding:'12px', background:'#f59e0b', color:'white', border:'none', borderRadius:'10px', fontSize:'15px', fontWeight:'bold', cursor:'pointer' }}>
              Reprendre la tournee
            </button>
          </div>
        )}

        <div style={{ marginBottom:'24px' }}>
          <label style={{ display:'block', fontSize:'18px', fontWeight:'600', color:'#374151', marginBottom:'10px' }}>Nom du chauffeur</label>
          <input type="text" value={nom}
            onChange={e => setNom(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && nom.trim() && onLogin(nom.trim())}
            placeholder="Votre nom"
            style={{ width:'100%', padding:'16px', fontSize:'20px', border:'2px solid #d1d5db', borderRadius:'12px', outline:'none', boxSizing:'border-box' }}
          />
        </div>
        <button onClick={() => nom.trim() && onLogin(nom.trim())}
          style={{ width:'100%', padding:'20px', background:'#2563eb', color:'white', border:'none', borderRadius:'12px', fontSize:'20px', fontWeight:'bold', cursor:'pointer', marginBottom:'12px' }}>
          DEMARRER
        </button>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px' }}>
          <button onClick={onConfig}
            style={{ padding:'14px', background:'#f3f4f6', color:'#6b7280', border:'2px solid #e5e7eb', borderRadius:'12px', fontSize:'14px', fontWeight:'600', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'6px' }}>
            <Settings size={16} /> Config
          </button>
          <button onClick={() => setApropos(true)}
            style={{ padding:'14px', background:'#eff6ff', color:'#2563eb', border:'2px solid #bfdbfe', borderRadius:'12px', fontSize:'14px', fontWeight:'600', cursor:'pointer' }}>
            A propos
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Écran d'import Excel ────────────────────────────────────────────────────
function VueImport({ onImport, onConfig }) {
  const cfg = chargerConfig();
  const [fichiersDrive, setFichiersDrive] = useState([]);
  const [chargement, setChargement]       = useState(false);
  const [erreur, setErreur]               = useState('');

  const DOSSIER_ID  = cfg.dossierDriveId || DOSSIERS_DRIVE[cfg.nomUnite] || '1SDu8PBRpHP7oTiT3U1CXXH82ScZYou7i';
  const API_KEY     = 'AIzaSyCt3Wzna095bA2G-JKLtlx8kTmOxOfTLSA';

  // Charger la liste des fichiers Excel depuis Google Drive au montage
  useEffect(() => {
    listerFichiersDrive();
  }, []);

  const listerFichiersDrive = async () => {
    setChargement(true);
    setErreur('');
    try {
      const url = `https://www.googleapis.com/drive/v3/files?q='${DOSSIER_ID}'+in+parents+and+mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'+and+trashed=false&key=${API_KEY}&fields=files(id,name,modifiedTime)&orderBy=modifiedTime+desc`;
      const res = await fetch(url);
      if (!res.ok) throw new Error('Erreur API Google Drive');
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      setFichiersDrive(data.files || []);
    } catch (err) {
      console.error('Erreur Drive:', err);
      setErreur('Impossible de charger les tournées depuis Drive. Vérifiez la connexion.');
    }
    setChargement(false);
  };

  const chargerDepuisDrive = async (fichier) => {
    setChargement(true);
    try {
      const url = `https://www.googleapis.com/drive/v3/files/${fichier.id}?alt=media&key=${API_KEY}`;
      const res = await fetch(url);
      if (!res.ok) throw new Error('Fichier inaccessible');
      const arrayBuffer = await res.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const clients = parserRowsExcel(rows);
      if (clients.length === 0) { alert("Aucun client trouvé dans ce fichier."); setChargement(false); return; }
      onImport({ id: `T-${new Date().toISOString().split('T')[0]}`, date: new Date().toISOString(), clients }, fichier.name.replace('.xlsx','').replace('.xls',''));
    } catch (err) {
      alert("Erreur de chargement : " + err.message);
      console.error(err);
    }
    setChargement(false);
  };

  const traiterFichierExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const clients = parserRowsExcel(rows);
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

  const formatDate = (iso) => {
    const d = new Date(iso);
    return d.toLocaleDateString('fr-FR') + ' ' + d.toLocaleTimeString('fr-FR', { hour:'2-digit', minute:'2-digit' });
  };

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'600px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'32px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>

        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'24px' }}>
          <h2 style={{ fontSize:'20px', fontWeight:'bold', color:'#1f2937', margin:0 }}>📥 Choisir une tournée</h2>
          <button onClick={listerFichiersDrive}
            style={{ background:'#dbeafe', border:'2px solid #93c5fd', borderRadius:'10px', padding:'8px 12px', cursor:'pointer', color:'#1e40af', fontSize:'14px', fontWeight:'600' }}>
            🔄 Actualiser
          </button>
        </div>

        {/* Tournées Google Drive */}
        <div style={{ marginBottom:'20px' }}>
          <div style={{ display:'flex', alignItems:'center', gap:'8px', marginBottom:'12px' }}>
            <span style={{ fontSize:'16px', fontWeight:'bold', color:'#374151' }}>📂 Tournées sur Google Drive</span>
            {chargement && <span style={{ fontSize:'13px', color:'#6b7280' }}>Chargement...</span>}
          </div>

          {erreur && (
            <div style={{ background:'#fef2f2', border:'2px solid #fecaca', borderRadius:'10px', padding:'12px', marginBottom:'12px' }}>
              <p style={{ color:'#dc2626', fontSize:'14px', margin:0 }}>⚠️ {erreur}</p>
            </div>
          )}

          {!chargement && !erreur && fichiersDrive.length === 0 && (
            <div style={{ background:'#f9fafb', border:'2px dashed #d1d5db', borderRadius:'10px', padding:'20px', textAlign:'center' }}>
              <p style={{ color:'#6b7280', fontSize:'14px', margin:0 }}>Aucun fichier Excel trouvé dans le dossier Drive.<br/>Déposez vos fichiers .xlsx dans le dossier BDL-Tournées.</p>
            </div>
          )}

          <div style={{ display:'grid', gap:'8px' }}>
            {fichiersDrive.map(fichier => (
              <button key={fichier.id} onClick={() => chargerDepuisDrive(fichier)}
                disabled={chargement}
                style={{ padding:'14px 16px', background:'#f0fdf4', border:'2px solid #86efac', borderRadius:'10px', cursor:'pointer', textAlign:'left', display:'flex', justifyContent:'space-between', alignItems:'center', opacity: chargement ? 0.6 : 1 }}>
                <span style={{ fontSize:'15px', fontWeight:'600', color:'#15803d' }}>📄 {fichier.name}</span>
                <span style={{ fontSize:'12px', color:'#6b7280' }}>{formatDate(fichier.modifiedTime)}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Séparateur */}
        <div style={{ display:'flex', alignItems:'center', gap:'12px', marginBottom:'16px' }}>
          <div style={{ flex:1, height:'1px', background:'#e5e7eb' }}></div>
          <span style={{ color:'#9ca3af', fontSize:'13px' }}>ou</span>
          <div style={{ flex:1, height:'1px', background:'#e5e7eb' }}></div>
        </div>

        {/* Import local */}
        <label htmlFor="file-upload" style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', border:'2px dashed #93c5fd', borderRadius:'12px', padding:'16px 20px', textAlign:'center', background:'#eff6ff', cursor:'pointer' }}>
          <FileText size={24} style={{ color:'#2563eb', flexShrink:0 }} />
          <span style={{ fontSize:'15px', fontWeight:'600', color:'#374151' }}>Importer un fichier local (.xlsx)</span>
          <input id="file-upload" type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:'none' }} />
        </label>

        {cfg.cheminExcel && (
          <div style={{ marginTop:'12px', background:'#f0f9ff', border:'2px solid #bae6fd', padding:'12px', borderRadius:'10px' }}>
            <p style={{ fontSize:'13px', color:'#0369a1', margin:0 }}>📁 Chemin configuré : <strong>{cfg.cheminExcel}</strong></p>
          </div>
        )}
      </div>
    </div>
  );
}


// ─── Sons feedback scan QR (Web Audio API — aucun fichier externe) ───────────
function _jouerNotes(notes) {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    notes.forEach(({ freq, debut, duree, type = 'sine' }) => {
      const osc  = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.type = type;
      osc.frequency.value = freq;
      gain.gain.setValueAtTime(0.25, ctx.currentTime + debut);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + debut + duree);
      osc.start(ctx.currentTime + debut);
      osc.stop(ctx.currentTime + debut + duree + 0.05);
    });
  } catch(e) { /* AudioContext non dispo */ }
}
const jouerSonOK     = () => _jouerNotes([{ freq:880, debut:0, duree:0.12 }, { freq:1320, debut:0.10, duree:0.18 }]);
const jouerSonErreur = () => _jouerNotes([{ freq:220, debut:0, duree:0.20, type:'sawtooth' }, { freq:180, debut:0.18, duree:0.28, type:'sawtooth' }]);
const jouerSonDejaLu = () => _jouerNotes([{ freq:660, debut:0, duree:0.09 }]);

// ─── Écran de récapitulatif + chargement cabris ──────────────────────────────
// ─── Scanner QR plein écran ───────────────────────────────────────────────────
// feedback : null | { type: 'ok'|'anomalie'|'dejaLu', ts: number }
function VueScanner({ onScan, onFermer, titre }) {
  const videoRef  = React.useRef(null);
  const streamRef = React.useRef(null);
  const timerRef  = React.useRef(null);
  const [erreur, setErreur]               = useState('');
  const [actif, setActif]                 = useState(false);
  const [dernierScan, setDernierScan]     = useState(null); // { val, heure, statut }
  const [saisieManuelle, setSaisieManuelle] = useState('');
  const [saisieOuverte, setSaisieOuverte] = useState(false);
  const derniereValRef  = React.useRef('');
  const dernierTempsRef = React.useRef(0);
  const supporte = 'BarcodeDetector' in window;

  // ── Démarrage caméra ──
  useEffect(() => {
    let mounted = true;
    if (!supporte) return;
    navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } } })
      .then(stream => {
        if (!mounted) { stream.getTracks().forEach(t => t.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          videoRef.current.play()
            .then(() => { if (mounted) setActif(true); })
            .catch(e => { if (mounted) setErreur('Impossible de démarrer la caméra : ' + e.message); });
        }
      }).catch(() => {
        if (mounted) setErreur("Accès caméra refusé. Autorisez l'accès dans les paramètres du navigateur.");
      });
    return () => {
      mounted = false;
      if (streamRef.current) streamRef.current.getTracks().forEach(t => t.stop());
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, [supporte]);

  // ── Détection QR toutes les 150 ms ──
  useEffect(() => {
    if (!actif) return;
    const detector = new window.BarcodeDetector({ formats: ['qr_code'] });
    timerRef.current = setInterval(async () => {
      const video = videoRef.current;
      if (!video || video.readyState < 2) return;
      try {
        const codes = await detector.detect(video);
        if (codes.length > 0) {
          const val = codes[0].rawValue.trim();
          const now = Date.now();
          if (val && (val !== derniereValRef.current || now - dernierTempsRef.current > 2000)) {
            derniereValRef.current  = val;
            dernierTempsRef.current = now;
            const result = onScan(val);
            if (result === 'ok')       jouerSonOK();
            else if (result === 'anomalie') jouerSonErreur();
            else if (result === 'dejaLu')  jouerSonDejaLu();
            setDernierScan({ val, heure: new Date().toLocaleTimeString('fr-FR'), statut: result || null });
          }
        }
      } catch {}
    }, 150);
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [actif, onScan]);

  const validerSaisie = () => {
    const val = saisieManuelle.trim();
    if (!val) return;
    const result = onScan(val);
    if (result === 'ok')       jouerSonOK();
    else if (result === 'anomalie') jouerSonErreur();
    else if (result === 'dejaLu')  jouerSonDejaLu();
    setDernierScan({ val, heure: new Date().toLocaleTimeString('fr-FR'), statut: result || null });
    setSaisieManuelle('');
  };

  // Couleur et icône selon statut
  const couleurStatut = (statut) => {
    if (statut === 'ok')       return '#86efac';
    if (statut === 'anomalie') return '#fca5a5';
    if (statut === 'dejaLu')   return '#fde68a';
    return '#94a3b8';
  };
  const iconeStatut = (statut) => {
    if (statut === 'ok')       return '✅';
    if (statut === 'anomalie') return '❌';
    if (statut === 'dejaLu')   return '⚠️';
    return '⏳';
  };
  const libelleStatut = (statut, val) => {
    if (statut === 'ok')       return `✅ ${val} — validé`;
    if (statut === 'anomalie') return `❌ ${val} — inconnu`;
    if (statut === 'dejaLu')   return `⚠️ ${val} — Déjà lu`;
    return `⏳ ${val} — en attente...`;
  };

  return (
    <div style={{ position:'fixed', inset:0, background:'#0f172a', zIndex:9999, display:'flex', flexDirection:'column' }}>

      {/* Header */}
      <div style={{ padding:'14px 18px', background:'rgba(0,0,0,0.85)', display:'flex', justifyContent:'space-between', alignItems:'center', flexShrink:0 }}>
        <span style={{ color:'white', fontSize:'17px', fontWeight:'bold' }}>📷 {titre}</span>
        <button onClick={onFermer} style={{ background:'#ef4444', color:'white', border:'none', borderRadius:'8px', padding:'10px 18px', fontSize:'15px', fontWeight:'bold', cursor:'pointer' }}>
          ✕ Fermer
        </button>
      </div>

      {/* Corps : caméra ou mode saisie manuelle forcée */}
      {!supporte || erreur ? (
        /* ── Fallback : pas de caméra ── */
        <div style={{ flex:1, display:'flex', alignItems:'center', justifyContent:'center', padding:'24px' }}>
          <div style={{ background:'#1e293b', borderRadius:'16px', padding:'32px', textAlign:'center', maxWidth:'340px', width:'100%' }}>
            <p style={{ fontSize:'48px', margin:'0 0 12px' }}>⌨️</p>
            {erreur && <p style={{ color:'#fca5a5', fontSize:'14px', marginBottom:'16px', lineHeight:'1.6' }}>{erreur}</p>}
            {!supporte && <p style={{ color:'#fde68a', fontSize:'14px', marginBottom:'16px', lineHeight:'1.6' }}>
              Scan automatique non disponible.<br/>Saisissez l'identifiant manuellement.
            </p>}
            <input type="text" value={saisieManuelle}
              onChange={e => setSaisieManuelle(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && validerSaisie()}
              placeholder="Ex: 12345"
              style={{ width:'100%', padding:'16px', fontSize:'20px', textAlign:'center', border:'2px solid #475569', borderRadius:'10px', outline:'none', boxSizing:'border-box', background:'#0f172a', color:'white', marginBottom:'12px' }}
              autoFocus
            />
            <button onClick={validerSaisie} disabled={!saisieManuelle.trim()}
              style={{ width:'100%', padding:'14px', background: saisieManuelle.trim() ? '#16a34a' : '#374151', color:'white', border:'none', borderRadius:'10px', fontSize:'16px', fontWeight:'bold', cursor: saisieManuelle.trim() ? 'pointer' : 'not-allowed', marginBottom:'12px' }}>
              ✅ Valider
            </button>
            {dernierScan && (
              <p style={{ color: couleurStatut(dernierScan.statut), fontSize:'14px', fontWeight:'600', margin:0 }}>
                {libelleStatut(dernierScan.statut, dernierScan.val)}
              </p>
            )}
          </div>
        </div>
      ) : (
        /* ── Mode caméra ── */
        <>
          <div style={{ flex:1, position:'relative', overflow:'hidden' }}>
            <video ref={videoRef} playsInline muted style={{ width:'100%', height:'100%', objectFit:'cover' }} />
            {/* Cadre viseur — change couleur selon dernier statut */}
            <div style={{ position:'absolute', inset:0, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'none' }}>
              <div style={{
                width:'220px', height:'220px', borderRadius:'16px',
                border: `3px solid ${dernierScan ? couleurStatut(dernierScan.statut) : '#22c55e'}`,
                boxShadow:'0 0 0 9999px rgba(0,0,0,0.45)',
                transition:'border-color 0.3s'
              }} />
            </div>
          </div>

          {/* Pied : saisie manuelle + statut dernier scan */}
          <div style={{ background:'rgba(0,0,0,0.88)', flexShrink:0 }}>

            {/* Zone saisie manuelle (toggle) */}
            {saisieOuverte && (
              <div style={{ padding:'10px 14px', borderBottom:'1px solid rgba(255,255,255,0.1)', display:'flex', gap:'8px', alignItems:'center' }}>
                <input type="text" value={saisieManuelle}
                  onChange={e => setSaisieManuelle(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && validerSaisie()}
                  placeholder="ID cabri (ex: 12345)"
                  style={{ flex:1, padding:'11px 14px', fontSize:'17px', textAlign:'center', border:'2px solid #475569', borderRadius:'8px', outline:'none', background:'#1e293b', color:'white', boxSizing:'border-box' }}
                  autoFocus
                />
                <button onClick={validerSaisie} disabled={!saisieManuelle.trim()}
                  style={{ padding:'11px 16px', background: saisieManuelle.trim() ? '#16a34a' : '#374151', color:'white', border:'none', borderRadius:'8px', fontSize:'16px', fontWeight:'bold', cursor: saisieManuelle.trim() ? 'pointer' : 'not-allowed' }}>
                  ✅
                </button>
                <button onClick={() => { setSaisieOuverte(false); setSaisieManuelle(''); }}
                  style={{ padding:'11px 13px', background:'#374151', color:'white', border:'none', borderRadius:'8px', fontSize:'16px', cursor:'pointer' }}>
                  ✕
                </button>
              </div>
            )}

            {/* Ligne statut + bouton saisie */}
            <div style={{ padding:'12px 16px', display:'flex', justifyContent:'space-between', alignItems:'center', gap:'10px' }}>
              <div style={{ flex:1 }}>
                {dernierScan ? (
                  <p style={{ margin:0, fontSize:'14px', fontWeight:'700', color: couleurStatut(dernierScan.statut) }}>
                    {libelleStatut(dernierScan.statut, dernierScan.val)}
                    <span style={{ fontWeight:'400', color:'#94a3b8', marginLeft:'8px', fontSize:'12px' }}>{dernierScan.heure}</span>
                  </p>
                ) : (
                  <p style={{ margin:0, color:'#94a3b8', fontSize:'13px' }}>Centrez le QR code dans le cadre vert</p>
                )}
              </div>
              <button
                onClick={() => setSaisieOuverte(v => !v)}
                title="Saisie manuelle"
                style={{ padding:'10px 14px', background: saisieOuverte ? '#2563eb' : '#475569', color:'white', border:'none', borderRadius:'8px', fontSize:'15px', cursor:'pointer', flexShrink:0 }}>
                ✏️
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ─── Modal anomalie QR ────────────────────────────────────────────────────────
function ModalAnomalieQR({ idCabri, message, onConfirmer, onAnnuler }) {
  const [commentaire, setCommentaire] = useState('');
  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.75)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:10000, padding:'20px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'28px', maxWidth:'400px', width:'100%', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign:'center', marginBottom:'20px' }}>
          <p style={{ fontSize:'48px', margin:0 }}>⚠️</p>
          <h3 style={{ fontSize:'20px', fontWeight:'bold', color:'#1f2937', margin:'8px 0 0' }}>Anomalie détectée</h3>
        </div>
        <div style={{ background:'#fef2f2', border:'2px solid #fecaca', borderRadius:'12px', padding:'16px', marginBottom:'16px' }}>
          <p style={{ fontWeight:'bold', color:'#dc2626', fontSize:'18px', margin:'0 0 6px', fontFamily:'monospace' }}>{idCabri}</p>
          <p style={{ color:'#7f1d1d', fontSize:'14px', margin:0, lineHeight:'1.5' }}>{message}</p>
        </div>
        <label style={{ display:'block', fontWeight:'600', color:'#374151', marginBottom:'8px', fontSize:'15px' }}>
          Commentaire obligatoire :
        </label>
        <textarea
          value={commentaire} onChange={e => setCommentaire(e.target.value)}
          rows={3} placeholder="Décrivez l'anomalie..."
          autoFocus
          style={{ width:'100%', padding:'12px', fontSize:'15px', border:'2px solid #fca5a5', borderRadius:'10px', outline:'none', resize:'vertical', boxSizing:'border-box', marginBottom:'16px' }}
        />
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'10px' }}>
          <button onClick={onAnnuler}
            style={{ padding:'14px', background:'#f3f4f6', color:'#374151', border:'2px solid #e5e7eb', borderRadius:'12px', fontSize:'15px', fontWeight:'bold', cursor:'pointer' }}>
            ← Rescanner
          </button>
          <button onClick={() => { if (commentaire.trim()) onConfirmer(idCabri, commentaire.trim()); }}
            disabled={!commentaire.trim()}
            style={{ padding:'14px', background: commentaire.trim() ? '#f59e0b' : '#9ca3af', color:'white', border:'none', borderRadius:'12px', fontSize:'15px', fontWeight:'bold', cursor: commentaire.trim() ? 'pointer' : 'not-allowed' }}>
            Signaler →
          </button>
        </div>
      </div>
    </div>
  );
}

function VueRecap({ tournee, onDemarrer, onRetour }) {
  const [ordre, setOrdre] = useState(tournee.clients.map((_, i) => i));
  const [chargementDemarre, setChargementDemarre] = useState(false);
  const [heureChargement, setHeureChargement] = useState(null);

  // ── État QR scan chargement ────────────────────────────────────────────────
  const [cabrisScannés, setCabrisScannés]             = useState([]);     // IDs validés
  const [anomaliesChargement, setAnomaliesChargement] = useState([]);     // anomalies signalées
  const [scannerOuvert, setScannerOuvert]             = useState(false);
  const [anomalieEnCours, setAnomalieEnCours]         = useState(null);   // { id }
  const [chargementConfirme, setChargementConfirme]   = useState(false);  // mode sans QR

  const clientsOrdres = ordre.map(i => tournee.clients[i]);
  const totalCabris = tournee.clients.reduce((t, c) => t + c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0), 0);

  // Liste exhaustive des cabris attendus pour cette tournée
  const tousLesCabris = React.useMemo(() => {
    const liste = [];
    tournee.clients.forEach(c => {
      c.services.forEach(s => {
        (s.cabrisIds || []).forEach(id => {
          liste.push({ id, clientId: c.id, clientNom: c.nom, serviceId: s.id, serviceNom: s.nom });
        });
      });
    });
    return liste;
  }, [tournee]);

  const modeQR = tousLesCabris.length > 0;
  const totalAttendu   = tousLesCabris.length;
  const totalTraite    = cabrisScannés.length + anomaliesChargement.length;
  const toutTraite     = totalAttendu === 0 || totalTraite >= totalAttendu;
  // Départ autorisé : en mode QR au moins 1 cabri traité, sinon checkbox manuelle
  const departAutorise = modeQR ? totalTraite > 0 : chargementConfirme;

  const handleScanChargement = React.useCallback((id) => {
    if (cabrisScannés.includes(id) || anomaliesChargement.some(a => a.id === id)) {
      return 'dejaLu';
    }
    const cabriInfo = tousLesCabris.find(c => c.id === id);
    if (cabriInfo) {
      const newScannés = [...cabrisScannés, id];
      setCabrisScannés(newScannés);
      // Auto-fermeture quand tous les cabris sont scannés
      if (newScannés.length + anomaliesChargement.length >= tousLesCabris.length) {
        setScannerOuvert(false);
      }
      return 'ok';
    } else {
      setScannerOuvert(false);
      setAnomalieEnCours({ id });
      return 'anomalie';
    }
  }, [cabrisScannés, anomaliesChargement, tousLesCabris]);

  const confirmerAnomalie = (id, commentaire) => {
    setAnomaliesChargement(prev => [...prev, { id, commentaire }]);
    setAnomalieEnCours(null);
    setScannerOuvert(true);
  };

  const confirmerDepart = () => {
    if (modeQR && !toutTraite) {
      const manquants = totalAttendu - totalTraite;
      const ok = window.confirm(`⚠️ Attention : ${manquants} cabri(s) non scanné(s).\n\nConfirmer le départ quand même ?`);
      if (!ok) return;
    }
    const cabrisChargement = modeQR ? { scannés: cabrisScannés, anomalies: anomaliesChargement } : null;
    onDemarrer(clientsOrdres, heureChargement.toISOString(), cabrisChargement);
  };

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
    <>
      {/* Scanner QR plein écran */}
      {scannerOuvert && (
        <VueScanner
          titre="Scanner cabri — Chargement"
          onScan={handleScanChargement}
          onFermer={() => setScannerOuvert(false)}
        />
      )}

      {/* Modal anomalie */}
      {anomalieEnCours && (
        <ModalAnomalieQR
          idCabri={anomalieEnCours.id}
          message="Ce cabri n'appartient pas à cette tournée."
          onConfirmer={confirmerAnomalie}
          onAnnuler={() => { setAnomalieEnCours(null); setScannerOuvert(true); }}
        />
      )}

    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'8px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'16px', padding:'16px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'12px' }}>
          <h2 style={{ fontSize:'20px', fontWeight:'bold', margin:0 }}>
            <CheckCircle size={22} style={{ color:'#16a34a', verticalAlign:'middle', marginRight:'8px' }} />
            Clients à livrer
          </h2>
          <button onClick={onRetour}
            style={{ padding:'8px 12px', background:'#f3f4f6', color:'#374151', border:'2px solid #e5e7eb', borderRadius:'8px', fontSize:'13px', fontWeight:'600', cursor:'pointer' }}>
            ← Changer
          </button>
        </div>

        {/* Compteurs */}
        <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'8px', marginBottom:'12px' }}>
          {[
            { label:'Clients',      value: tournee.clients.length,                             color:'#2563eb', bg:'#eff6ff' },
            { label:'Total Cabris', value: totalCabris,                                        color:'#16a34a', bg:'#f0fdf4' },
            { label:'Date',         value: new Date(tournee.date).toLocaleDateString('fr-FR'), color:'#7c3aed', bg:'#faf5ff', small: true },
          ].map(({ label, value, color, bg, small }) => (
            <div key={label} style={{ background:bg, borderRadius:'10px', padding:'10px', textAlign:'center' }}>
              <p style={{ color:'#6b7280', marginBottom:'3px', fontSize:'12px' }}>{label}</p>
              <p style={{ fontSize: small ? '13px' : '20px', fontWeight:'bold', color, margin:0, wordBreak:'break-word', lineHeight:1.2 }}>{value}</p>
            </div>
          ))}
        </div>

        {/* Ordre de livraison */}
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'6px' }}>
          <p style={{ fontWeight:'bold', fontSize:'15px', margin:0 }}>🗂️ Ordre de livraison</p>
          <p style={{ fontSize:'11px', color:'#6b7280', margin:0 }}>↑↓ pour réorganiser</p>
        </div>
        <div style={{ marginBottom:'12px' }}>
          {clientsOrdres.map((c, pos) => (
            <div key={c.id} style={{
              background: pos === 0 ? '#f0fdf4' : '#f9fafb',
              border: pos === 0 ? '2px solid #86efac' : '1px solid #e5e7eb',
              borderRadius:'10px', padding:'7px 10px', marginBottom:'4px',
              display:'flex', alignItems:'center', gap:'8px'
            }}>
              <div style={{
                minWidth:'26px', height:'26px', borderRadius:'50%', display:'flex', alignItems:'center', justifyContent:'center',
                background: pos === 0 ? '#16a34a' : '#e5e7eb',
                color: pos === 0 ? 'white' : '#374151', fontWeight:'bold', fontSize:'13px', flexShrink:0
              }}>{pos + 1}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <p style={{ fontWeight:'bold', fontSize:'14px', margin:0, wordBreak:'break-word', lineHeight:'1.3' }}>
                  {c.nom}
                </p>
                <p style={{ color:'#6b7280', fontSize:'11px', margin:0 }}>
                  {c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0)} cabris
                </p>
              </div>
              <div style={{ display:'flex', flexDirection:'column', gap:'3px', flexShrink:0 }}>
                <BtnOrdre onClick={() => monter(pos)}    label="↑" disabled={pos === 0} />
                <BtnOrdre onClick={() => descendre(pos)} label="↓" disabled={pos === ordre.length - 1} />
              </div>
            </div>
          ))}
        </div>

        {/* ── Phase chargement ─────────────────────────────────────────────── */}
        {!chargementDemarre ? (
          <button onClick={demarrerChargement}
            style={{ width:'100%', padding:'16px', background:'#f59e0b', color:'white', border:'none', borderRadius:'12px', fontSize:'18px', fontWeight:'bold', cursor:'pointer' }}>
            📦 Démarrer le chargement
          </button>
        ) : (
          <div>
            {/* Heure chargement */}
            <div style={{ background:'#fefce8', border:'2px solid #fde68a', borderRadius:'10px', padding:'8px 12px', marginBottom:'10px', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
              <p style={{ fontWeight:'bold', color:'#92400e', fontSize:'14px', margin:0 }}>⏱️ Chargement en cours...</p>
              <p style={{ color:'#b45309', fontSize:'13px', margin:0 }}>
                Début : {heureChargement.toLocaleTimeString('fr-FR')}
              </p>
            </div>

            {/* ── SECTION SCAN QR (uniquement si mode QR activé) ──────────── */}
            {modeQR && (
              <div style={{ background:'#f0f9ff', border:'2px solid #bae6fd', borderRadius:'12px', padding:'12px', marginBottom:'10px' }}>
                {/* En-tête scan */}
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'10px' }}>
                  <p style={{ fontWeight:'bold', fontSize:'15px', color:'#0369a1', margin:0 }}>
                    📦 Scan des cabris
                  </p>
                  <span style={{
                    background: toutTraite ? '#dcfce7' : '#fef3c7',
                    color: toutTraite ? '#15803d' : '#92400e',
                    fontWeight:'bold', fontSize:'14px', padding:'4px 12px', borderRadius:'999px'
                  }}>
                    {totalTraite}/{totalAttendu}
                  </span>
                </div>

                {/* Barre de progression */}
                <div style={{ background:'#e0f2fe', borderRadius:'999px', height:'10px', marginBottom:'16px', overflow:'hidden' }}>
                  <div style={{
                    background: toutTraite ? '#16a34a' : '#0284c7',
                    height:'100%', borderRadius:'999px',
                    width: totalAttendu > 0 ? `${Math.round(totalTraite / totalAttendu * 100)}%` : '0%',
                    transition:'width 0.3s'
                  }} />
                </div>

                {/* Liste des cabris par client */}
                {clientsOrdres.map(c => {
                  const serviceAvecIDs = c.services.filter(s => s.cabrisIds?.length > 0);
                  if (serviceAvecIDs.length === 0) return null;
                  return (
                    <div key={c.id} style={{ marginBottom:'12px', background:'white', borderRadius:'10px', padding:'12px', border:'1px solid #e0f2fe' }}>
                      <p style={{ fontWeight:'bold', fontSize:'14px', color:'#0369a1', margin:'0 0 8px' }}>{c.nom}</p>
                      {serviceAvecIDs.map(s => (
                        <div key={s.id} style={{ marginBottom:'6px' }}>
                          <p style={{ fontSize:'12px', color:'#6b7280', margin:'0 0 4px' }}>{s.nom}</p>
                          <div style={{ display:'flex', flexWrap:'wrap', gap:'6px' }}>
                            {s.cabrisIds.map(id => {
                              const scanne = cabrisScannés.includes(id);
                              const anomalie = anomaliesChargement.some(a => a.id === id);
                              return (
                                <span key={id} style={{
                                  padding:'4px 10px', borderRadius:'8px', fontSize:'12px', fontWeight:'bold', fontFamily:'monospace',
                                  background: scanne ? '#dcfce7' : anomalie ? '#fef3c7' : '#f1f5f9',
                                  color: scanne ? '#15803d' : anomalie ? '#92400e' : '#64748b',
                                  border: `1px solid ${scanne ? '#86efac' : anomalie ? '#fde68a' : '#e2e8f0'}`
                                }}>
                                  {scanne ? '✅' : anomalie ? '⚠️' : '⏳'} {id}
                                </span>
                              );
                            })}
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                })}

                {/* Anomalies */}
                {anomaliesChargement.length > 0 && (
                  <div style={{ background:'#fefce8', border:'1px solid #fde68a', borderRadius:'10px', padding:'12px', marginBottom:'12px' }}>
                    <p style={{ fontWeight:'bold', color:'#92400e', fontSize:'13px', margin:'0 0 6px' }}>⚠️ {anomaliesChargement.length} anomalie(s) signalée(s)</p>
                    {anomaliesChargement.map((a, i) => (
                      <p key={i} style={{ fontSize:'12px', color:'#78350f', margin:'2px 0', fontFamily:'monospace' }}>
                        {a.id} — "{a.commentaire}"
                      </p>
                    ))}
                  </div>
                )}

                {/* Bouton scanner */}
                <button onClick={() => setScannerOuvert(true)} disabled={toutTraite}
                  style={{ width:'100%', padding:'16px', background: toutTraite ? '#6b7280' : '#0284c7', color:'white', border:'none', borderRadius:'12px', fontSize:'16px', fontWeight:'bold', cursor: toutTraite ? 'default' : 'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', opacity: toutTraite ? 0.7 : 1 }}>
                  {toutTraite ? '✅ Tous les cabris scannés' : '📱 Scanner un cabri'}
                </button>
              </div>
            )}

            {/* Confirmation chargement (mode sans QR uniquement) */}
            {!modeQR && (
              <label style={{ display:'flex', alignItems:'center', gap:'12px', background:'#f0fdf4', border:`2px solid ${chargementConfirme ? '#86efac' : '#fca5a5'}`, borderRadius:'12px', padding:'14px', marginBottom:'10px', cursor:'pointer' }}>
                <input type="checkbox" checked={chargementConfirme} onChange={e => setChargementConfirme(e.target.checked)}
                  style={{ width:'22px', height:'22px', cursor:'pointer', accentColor:'#16a34a' }} />
                <span style={{ fontWeight:'bold', fontSize:'15px', color: chargementConfirme ? '#15803d' : '#dc2626' }}>
                  {chargementConfirme ? '✅ Camion chargé confirmé' : '⚠️ Confirmer que le camion est chargé'}
                </span>
              </label>
            )}

            {/* Message blocage QR */}
            {modeQR && totalTraite === 0 && (
              <div style={{ background:'#fef2f2', border:'2px solid #fca5a5', borderRadius:'10px', padding:'10px 14px', marginBottom:'10px', textAlign:'center' }}>
                <p style={{ color:'#dc2626', fontWeight:'bold', fontSize:'14px', margin:0 }}>
                  🚫 Scanner au moins un cabri avant de partir
                </p>
              </div>
            )}

            {/* Bouton DÉPART */}
            <button onClick={confirmerDepart} disabled={!departAutorise}
              style={{ width:'100%', padding:'16px', background: !departAutorise ? '#9ca3af' : (modeQR && !toutTraite ? '#f59e0b' : '#16a34a'), color:'white', border:'none', borderRadius:'12px', fontSize:'18px', fontWeight:'bold', cursor: departAutorise ? 'pointer' : 'not-allowed' }}>
              <Clock size={22} style={{ verticalAlign:'middle', marginRight:'8px' }} />
              {modeQR && !toutTraite ? `🚚 DÉPART (${totalAttendu - totalTraite} manquant(s))` : '🚚 DÉPART DE L\'UNITÉ'}
            </button>
          </div>
        )}
      </div>
    </div>
    </>
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
          <p style={{ color:'#6b7280', fontSize:'14px', marginBottom:'6px' }}>Services a livrer</p>
          <p style={{ fontSize:'22px', fontWeight:'bold', color:'#2563eb', margin:0 }}>
            {client.services.length} service(s) — {client.services.reduce((s, sv) => s + sv.cabrisPrevu, 0)} cabris
          </p>
        </div>


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
function VueClient({ client, onDepart, tournee }) {
  const [services, setServices]       = useState(() => client.services.map(s => ({ ...s })));
  const [cabrisReprisTotal, setCabrisReprisTotal] = useState('');
  const [commentaire, setCommentaire] = useState('');
  const [photos, setPhotos]               = useState([]);
  const [photoAgrandie, setPhotoAgrandie] = useState(null);
  const refCamera  = React.useRef(null);
  const refGalerie = React.useRef(null);

  // ── État QR scan livraison ─────────────────────────────────────────────────
  const [cabrisScannésLiv, setCabrisScannésLiv]     = useState({});   // { serviceId: [ids] }
  const [anomaliesLiv, setAnomaliesLiv]             = useState([]);   // anomalies livraison
  const [scannerLivOuvert, setScannerLivOuvert]     = useState(false);
  const [anomalieLivEnCours, setAnomalieLivEnCours] = useState(null); // { id, message }

  const modeQRLiv = client.services.some(s => s.cabrisIds?.length > 0);

  const totalCabrisAttendus = client.services.reduce((t, s) => t + (s.cabrisIds?.length || 0), 0);
  const totalCabrisTraites  = Object.values(cabrisScannésLiv).flat().length + anomaliesLiv.length;
  const toutTraiteLiv       = totalCabrisAttendus === 0 || totalCabrisTraites >= totalCabrisAttendus;

  const handleScanLivraison = React.useCallback((id) => {
    const dejaScanné = Object.values(cabrisScannésLiv).flat().includes(id) || anomaliesLiv.some(a => a.id === id);
    if (dejaScanné) return 'dejaLu';

    let serviceMatched = null;
    client.services.forEach(s => {
      if ((s.cabrisIds || []).includes(id)) serviceMatched = s;
    });

    if (serviceMatched) {
      const newServiceScannés = [...(cabrisScannésLiv[serviceMatched.id] || []), id];
      const newCabrisScannésLiv = { ...cabrisScannésLiv, [serviceMatched.id]: newServiceScannés };
      setCabrisScannésLiv(newCabrisScannésLiv);
      // Auto-fermeture quand tous les cabris sont scannés
      const newTotal = Object.values(newCabrisScannésLiv).flat().length + anomaliesLiv.length;
      if (newTotal >= totalCabrisAttendus) setScannerLivOuvert(false);
      return 'ok';
    } else {
      setScannerLivOuvert(false);
      let autreClient = null;
      if (tournee) {
        tournee.clients.forEach(c => {
          if (c.id !== client.id) {
            c.services.forEach(s => {
              if ((s.cabrisIds || []).includes(id)) autreClient = c;
            });
          }
        });
      }
      const msg = autreClient
        ? `Ce cabri appartient au client "${autreClient.nom}", pas à "${client.nom}".`
        : `Ce cabri n'est pas prévu pour ${client.nom}.`;
      setAnomalieLivEnCours({ id, message: msg });
      return 'anomalie';
    }
  }, [cabrisScannésLiv, anomaliesLiv, client, tournee, totalCabrisAttendus]);

  const confirmerAnomalieLiv = (id, commentaire) => {
    setAnomaliesLiv(prev => [...prev, { id, commentaire, type: 'mauvais_cabri' }]);
    setAnomalieLivEnCours(null);
    setScannerLivOuvert(true);
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
  const btnStyle = (bg) => ({ width:'100%', padding:'12px', background:bg, color:'white', border:'none', borderRadius:'10px', fontSize:'15px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px' });

  return (
    <>
      {/* Scanner QR livraison */}
      {scannerLivOuvert && (
        <VueScanner
          titre={`Scanner cabris — ${client.nom}`}
          onScan={handleScanLivraison}
          onFermer={() => setScannerLivOuvert(false)}
        />
      )}
      {/* Modal anomalie livraison */}
      {anomalieLivEnCours && (
        <ModalAnomalieQR
          idCabri={anomalieLivEnCours.id}
          message={anomalieLivEnCours.message}
          onConfirmer={confirmerAnomalieLiv}
          onAnnuler={() => { setAnomalieLivEnCours(null); setScannerLivOuvert(true); }}
        />
      )}

    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'8px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'16px', padding:'14px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'20px', fontWeight:'bold', color:'#1f2937', marginBottom:'8px' }}>{client.nom}</h2>

        {/* Contact */}
        <div style={{ background:'#f0fdf4', borderRadius:'10px', padding:'8px 12px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'10px' }}>
          <span style={{ fontWeight:'600', color:'#15803d', fontSize:'12px', flexShrink:0 }}>📞</span>
          <span style={{ fontSize:'16px', fontWeight:'bold' }}>{client.telephone || 'Non renseigne'}</span>
        </div>

        {client.noteTournee && (
          <div style={{ background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'10px', padding:'8px 12px', marginBottom:'8px' }}>
            <p style={{ fontWeight:'700', color:'#92400e', marginBottom:'2px', fontSize:'12px' }}>📝 Note de tournée</p>
            <p style={{ color:'#374151', margin:0, fontSize:'13px', lineHeight:'1.5', whiteSpace:'pre-wrap' }}>{client.noteTournee}</p>
          </div>
        )}
        <h3 style={{ fontSize:'16px', fontWeight:'bold', marginBottom:'6px' }}>Services à livrer</h3>
        {/* Affichage lecture seule par service (avec IDs si mode QR) */}
        {services.map((sv) => (
          <div key={sv.id} style={{ background:'#f9fafb', borderRadius:'8px', padding:'8px 10px', marginBottom:'4px' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom: sv.cabrisIds?.length > 0 ? '6px' : '0' }}>
              <span style={{ fontWeight:'600', fontSize:'14px', color:'#374151' }}>{sv.nom}</span>
              <div style={{ background:'#dbeafe', border:'1px solid #93c5fd', borderRadius:'6px', padding:'3px 10px', textAlign:'center' }}>
                <span style={{ fontSize:'16px', fontWeight:'bold', color:'#1d4ed8' }}>{sv.cabrisPrevu}</span>
                <span style={{ fontSize:'10px', color:'#6b7280', marginLeft:'3px' }}>cabris</span>
              </div>
            </div>
            {/* IDs cabris avec statut scan */}
            {(sv.cabrisIds?.length > 0) && (
              <div style={{ display:'flex', flexWrap:'wrap', gap:'6px' }}>
                {sv.cabrisIds.map(id => {
                  const scannéOk = (cabrisScannésLiv[sv.id] || []).includes(id);
                  return (
                    <span key={id} style={{
                      padding:'3px 10px', borderRadius:'8px', fontSize:'12px', fontWeight:'bold', fontFamily:'monospace',
                      background: scannéOk ? '#dcfce7' : '#f1f5f9',
                      color: scannéOk ? '#15803d' : '#64748b',
                      border: `1px solid ${scannéOk ? '#86efac' : '#e2e8f0'}`
                    }}>
                      {scannéOk ? '✅' : '⏳'} {id}
                    </span>
                  );
                })}
              </div>
            )}
          </div>
        ))}

        {/* ── SECTION SCAN QR livraison ────────────────────────────────────── */}
        {modeQRLiv && (
          <div style={{ background:'#f0f9ff', border:'2px solid #bae6fd', borderRadius:'12px', padding:'10px', marginTop:'8px', marginBottom:'8px' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'6px' }}>
              <p style={{ fontWeight:'bold', fontSize:'14px', color:'#0369a1', margin:0 }}>📦 Scan cabris livrés</p>
              <span style={{
                background: toutTraiteLiv ? '#dcfce7' : '#fef3c7',
                color: toutTraiteLiv ? '#15803d' : '#92400e',
                fontWeight:'bold', fontSize:'13px', padding:'2px 8px', borderRadius:'999px'
              }}>
                {totalCabrisTraites}/{totalCabrisAttendus}
              </span>
            </div>
            {/* Barre progression */}
            <div style={{ background:'#e0f2fe', borderRadius:'999px', height:'6px', marginBottom:'8px', overflow:'hidden' }}>
              <div style={{
                background: toutTraiteLiv ? '#16a34a' : '#0284c7',
                height:'100%', borderRadius:'999px',
                width: totalCabrisAttendus > 0 ? `${Math.round(totalCabrisTraites / totalCabrisAttendus * 100)}%` : '0%',
                transition:'width 0.3s'
              }} />
            </div>
            {/* Anomalies livraison */}
            {anomaliesLiv.length > 0 && (
              <div style={{ background:'#fefce8', border:'1px solid #fde68a', borderRadius:'8px', padding:'10px', marginBottom:'10px' }}>
                <p style={{ fontWeight:'bold', color:'#92400e', fontSize:'13px', margin:'0 0 4px' }}>⚠️ {anomaliesLiv.length} anomalie(s)</p>
                {anomaliesLiv.map((a, i) => (
                  <p key={i} style={{ fontSize:'12px', color:'#78350f', margin:'2px 0', fontFamily:'monospace' }}>
                    {a.id} — "{a.commentaire}"
                  </p>
                ))}
              </div>
            )}
            <button onClick={() => setScannerLivOuvert(true)} disabled={toutTraiteLiv}
              style={{ width:'100%', padding:'14px', background: toutTraiteLiv ? '#6b7280' : '#0284c7', color:'white', border:'none', borderRadius:'12px', fontSize:'15px', fontWeight:'bold', cursor: toutTraiteLiv ? 'default' : 'pointer', opacity: toutTraiteLiv ? 0.7 : 1 }}>
              {toutTraiteLiv ? '✅ Tous les cabris scannés' : '📱 Scanner les cabris'}
            </button>
          </div>
        )}

        {/* Un seul champ : total cabris ramasses */}
        <div style={{ background:'#fef3c7', border: cabrisReprisTotal === '' ? '3px solid #ef4444' : '2px solid #fde68a', borderRadius:'10px', padding:'10px 12px', marginTop:'8px', marginBottom:'8px' }}>
          <p style={{ fontWeight:'bold', fontSize:'13px', color:'#92400e', margin:'0 0 6px' }}>
            Cabris ramassés — total (sale)
          </p>
          <input type="number" min="0"
            value={cabrisReprisTotal}
            onChange={e => setCabrisReprisTotal(e.target.value)}
            placeholder="0"
            style={{ width:'100%', padding:'8px', fontSize:'36px', fontWeight:'bold', textAlign:'center',
              border: cabrisReprisTotal === '' ? '2px solid #ef4444' : '2px solid #fde68a',
              borderRadius:'8px', outline:'none', boxSizing:'border-box', color:'#92400e' }}
          />
          {cabrisReprisTotal === '' && (
            <p style={{ color:'#ef4444', fontSize:'12px', margin:'4px 0 0', textAlign:'center' }}>
              Obligatoire — entrez 0 si aucun cabri ramassé
            </p>
          )}
        </div>

        {/* Commentaire de livraison */}
        <div style={{ marginBottom:'8px' }}>
          <p style={{ fontSize:'14px', fontWeight:'bold', marginBottom:'4px' }}>Commentaire</p>
          <textarea value={commentaire} onChange={e => setCommentaire(e.target.value)} rows={2}
            placeholder="Incident, observation, demande particulière..."
            style={{ width:'100%', padding:'8px', fontSize:'14px', border:'2px solid #d1d5db', borderRadius:'8px', outline:'none', resize:'vertical', boxSizing:'border-box' }}
          />
        </div>

        <div style={{ marginBottom:'10px', background:'#f0f9ff', border:'2px solid #bae6fd', borderRadius:'12px', padding:'10px 12px' }}>
          <p style={{ fontSize:'14px', fontWeight:'bold', marginBottom:'8px', color:'#0369a1' }}>
            📷 Photos {photos.length > 0 && `(${photos.length})`}
          </p>
          <input ref={refCamera}  type="file" accept="image/*" capture="environment" onChange={traiterFichiers} style={{ position:'absolute', width:'1px', height:'1px', opacity:0, pointerEvents:'none' }} />
          <input ref={refGalerie} type="file" accept="image/*" multiple onChange={traiterFichiers} style={{ position:'absolute', width:'1px', height:'1px', opacity:0, pointerEvents:'none' }} />
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px', marginBottom:'8px' }}>
            <button onClick={() => refCamera.current.click()}  style={btnStyle('#0284c7')}>📷 Photo</button>
            <button onClick={() => refGalerie.current.click()} style={btnStyle('#7c3aed')}>🖼️ Galerie</button>
          </div>
          {photos.length === 0 ? (
            <div style={{ textAlign:'center', padding:'12px', color:'#94a3b8', border:'2px dashed #cbd5e1', borderRadius:'8px', fontSize:'13px' }}>Aucune photo prise</div>
          ) : (
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px' }}>
              {photos.map(photo => (
                <div key={photo.id} style={{ position:'relative', borderRadius:'8px', overflow:'hidden', border:'2px solid #e2e8f0' }}>
                  <img src={photo.data} alt="livraison" onClick={() => setPhotoAgrandie(photo)}
                    style={{ width:'100%', height:'80px', objectFit:'cover', display:'block', cursor:'pointer' }} />
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

        <button
          onClick={() => {
            if (cabrisReprisTotal === '') {
              alert('Veuillez saisir le nombre de cabris ramasses (0 si aucun)');
              return;
            }
            // Vérification QR livraison
            if (modeQRLiv && !toutTraiteLiv) {
              const manquants = totalCabrisAttendus - totalCabrisTraites;
              const ok = window.confirm(`⚠️ Attention : ${manquants} cabri(s) non scanné(s) pour ${client.nom}.\n\nConfirmer la livraison quand même ?`);
              if (!ok) return;
            }
            onDepart(commentaire, services, photos, parseInt(cabrisReprisTotal) || 0, {
              cabrisScannés: cabrisScannésLiv,
              anomalies: anomaliesLiv
            });
          }}
          disabled={cabrisReprisTotal === ''}
          style={{ width:'100%', padding:'18px', background: cabrisReprisTotal === '' ? '#9ca3af' : '#16a34a', color:'white', border:'none', borderRadius:'12px', fontSize:'18px', fontWeight:'bold', cursor: cabrisReprisTotal === '' ? 'not-allowed' : 'pointer' }}>
          <ArrowRight size={22} style={{ verticalAlign:'middle', marginRight:'8px' }} />
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
    </>
  );
}

// ─── Écran fin de tournée (retour unité) ────────────────────────────────────
// ─── Écran retour à l'unité ────────────────────────────────────────────────
function VueRetourUnite({ onArriveeUnite, adresseUnite, logs }) {
  const ouvrirWaze = () => {
    const dest = adresseUnite || 'unité de départ';
    window.open(`https://waze.com/ul?q=${encodeURIComponent(dest)}&navigate=yes`, '_blank');
  };
  const dernierDepart = [...logs].reverse().find(l => l.type === 'DEPART_CLIENT');
  const refDepart = React.useRef(dernierDepart ? new Date(dernierDepart.timestamp) : new Date());
  const [elapsed, setElapsed] = useState(() => Math.floor((Date.now() - refDepart.current) / 1000));
  useEffect(() => {
    const t = setInterval(() => setElapsed(Math.floor((Date.now() - refDepart.current) / 1000)), 1000);
    return () => clearInterval(t);
  }, []);
  const fmt = (s) => `${String(Math.floor(s/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`;

  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#15803d,#166534)', display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'28px', maxWidth:'480px', width:'100%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <CheckCircle size={64} style={{ color:'#16a34a', margin:'0 auto 14px' }} />
        <h2 style={{ fontSize:'24px', fontWeight:'bold', color:'#1f2937', marginBottom:'6px' }}>
          Tous les clients sont livrés !
        </h2>
        <p style={{ color:'#6b7280', fontSize:'14px', marginBottom:'18px' }}>
          Retournez à l'unité pour décharger le linge sale.
        </p>

        {/* Chrono depuis la dernière livraison */}
        <div style={{ background:'#f0fdf4', border:'2px solid #86efac', borderRadius:'12px', padding:'12px', marginBottom:'16px' }}>
          <p style={{ color:'#15803d', fontSize:'12px', fontWeight:'600', margin:'0 0 4px' }}>⏱️ Temps depuis la dernière livraison</p>
          <p style={{ fontSize:'32px', fontWeight:'bold', color:'#166534', margin:0, fontFamily:'monospace' }}>{fmt(elapsed)}</p>
        </div>

        {adresseUnite && (
          <button onClick={ouvrirWaze}
            style={{ width:'100%', padding:'16px', background:'#2563eb', color:'white', border:'none', borderRadius:'12px', fontSize:'16px', fontWeight:'bold', cursor:'pointer', marginBottom:'10px', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
            <MapPin size={20} /> Retour à l'unité via Waze
          </button>
        )}

        <button onClick={onArriveeUnite}
          style={{ width:'100%', padding:'20px', background:'#16a34a', color:'white', border:'none', borderRadius:'14px', fontSize:'20px', fontWeight:'bold', cursor:'pointer' }}>
          🏠 Je suis arrivé à l'unité
        </button>
      </div>
    </div>
  );
}

// ─── Écran déchargement ───────────────────────────────────────────────────────
function VueDecharge({ onDechargeTerminee, logs }) {
  const logArrivee = logs.find(l => l.type === 'ARRIVEE_UNITE');
  const refArrivee = React.useRef(logArrivee ? new Date(logArrivee.timestamp) : new Date());
  const [elapsed, setElapsed] = useState(() => Math.floor((Date.now() - refArrivee.current) / 1000));
  useEffect(() => {
    const t = setInterval(() => setElapsed(Math.floor((Date.now() - refArrivee.current) / 1000)), 1000);
    return () => clearInterval(t);
  }, []);
  const fmt = (s) => `${String(Math.floor(s/60)).padStart(2,'0')}:${String(s%60).padStart(2,'0')}`;

  return (
    <div style={{ minHeight:'100vh', background:'linear-gradient(135deg,#92400e,#78350f)', display:'flex', alignItems:'center', justifyContent:'center', padding:'20px' }}>
      <div style={{ background:'white', borderRadius:'20px', padding:'28px', maxWidth:'480px', width:'100%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ fontSize:'56px', margin:'0 auto 14px' }}>🚚</div>
        <h2 style={{ fontSize:'24px', fontWeight:'bold', color:'#1f2937', marginBottom:'6px' }}>
          Déchargement du camion
        </h2>

        {/* Heure d'arrivée */}
        <div style={{ background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'12px', padding:'10px', marginBottom:'12px' }}>
          <p style={{ color:'#92400e', fontSize:'12px', fontWeight:'600', margin:'0 0 2px' }}>🏠 Arrivée à l'unité</p>
          <p style={{ fontSize:'22px', fontWeight:'bold', color:'#78350f', margin:0 }}>
            {refArrivee.current.toLocaleTimeString('fr-FR')}
          </p>
        </div>

        {/* Chrono déchargement */}
        <div style={{ background:'#fff7ed', border:'2px solid #fed7aa', borderRadius:'12px', padding:'14px', marginBottom:'18px' }}>
          <p style={{ color:'#c2410c', fontSize:'12px', fontWeight:'600', margin:'0 0 6px' }}>⏱️ Temps de déchargement</p>
          <p style={{ fontSize:'44px', fontWeight:'bold', color:'#9a3412', margin:0, fontFamily:'monospace' }}>{fmt(elapsed)}</p>
        </div>

        <div style={{ background:'#fef3c7', border:'1px solid #fde68a', borderRadius:'10px', padding:'10px', marginBottom:'16px' }}>
          <p style={{ color:'#92400e', fontSize:'13px', margin:0 }}>⛽ Pensez à faire le plein du camion !</p>
        </div>

        <button onClick={onDechargeTerminee}
          style={{ width:'100%', padding:'20px', background:'#dc2626', color:'white', border:'none', borderRadius:'14px', fontSize:'19px', fontWeight:'bold', cursor:'pointer' }}>
          ✅ Déchargement terminé — Clôturer
        </button>
      </div>
    </div>
  );
}

// ─── Écran rapport final ─────────────────────────────────────────────────────
function VueRapport({ tournee, clientData, logs, startTime, agentName, onNouvelle, nomTournee, cabrisChargement }) {
  const cfg = chargerConfig();
  // config est chargé depuis config.json + localStorage via le state

  const logDepart      = logs.find(l => l.type === 'DEPART_UNITE');
  const logChargement  = logs.find(l => l.type === 'DEBUT_CHARGEMENT');
  const logArrivee     = logs.find(l => l.type === 'ARRIVEE_UNITE');
  const logFin         = logs.find(l => l.type === 'FIN_DECHARGE');

  const dureeTotal = (startTime && logFin)
    ? Math.round((new Date(logFin.timestamp) - new Date(startTime)) / 60000) : 0;
  const dureeDecharge = (logArrivee && logFin)
    ? Math.round((new Date(logFin.timestamp) - new Date(logArrivee.timestamp)) / 60000) : null;

  const creerDocPDF = () => {
    const doc = new jsPDF();
    const pageW = 210; // largeur A4 mm
    const mL = 12;     // marge gauche
    const mR = 12;     // marge droite
    const contW = pageW - mL - mR; // largeur utile = 186mm

    // ── En-tête compact : logo + titre sur la même ligne ──
    try { doc.addImage(LOGO_BDL_PDF, 'JPEG', mL, 6, 22, 15); } catch(e) {}
    doc.setFontSize(14); doc.setFont(undefined, 'bold');
    doc.text('BDL-Livraison — Rapport de tournee', mL + 25, 12);
    doc.setFontSize(8); doc.setFont(undefined, 'normal');
    doc.setTextColor(100, 100, 100);

    // Infos tournée sur une seule ligne condensée
    const dateStr  = new Date(tournee.date).toLocaleDateString('fr-FR');
    const arrStr      = logChargement ? new Date(logChargement.timestamp).toLocaleTimeString('fr-FR') : '--';
    const depStr      = logDepart     ? new Date(logDepart.timestamp).toLocaleTimeString('fr-FR')     : '--';
    const arriveeStr  = logArrivee    ? new Date(logArrivee.timestamp).toLocaleTimeString('fr-FR')    : '--';
    const finStr      = logFin        ? new Date(logFin.timestamp).toLocaleTimeString('fr-FR')        : '--';
    const dureeStr    = `${Math.floor(dureeTotal/60)}h${String(dureeTotal%60).padStart(2,'0')}`;
    const dechargeStr = dureeDecharge !== null ? `${dureeDecharge}min` : '--';
    doc.text(`${nomTournee || ''}  |  ${agentName}  |  ${dateStr}  |  Chgt: ${arrStr}  Dep: ${depStr}  Arr: ${arriveeStr}  Fin: ${finStr}  Decharge: ${dechargeStr}  Duree: ${dureeStr}`, mL + 25, 18);
    doc.setTextColor(0, 0, 0);

    // Ligne séparatrice
    doc.setDrawColor(37, 99, 235); doc.setLineWidth(0.6);
    doc.line(mL, 23, pageW - mR, 23);

    let y = 28;

    tournee.clients.forEach((client, i) => {
      const d = clientData[client.id];
      if (!d) return;
      if (y > 265) { doc.addPage(); y = 15; }

      // ── Nom client (fond bleu très clair) ──
      doc.setFillColor(239, 246, 255);
      doc.rect(mL, y, contW, 7, 'F');
      doc.setFontSize(10); doc.setFont(undefined, 'bold');
      doc.text(`${i+1}. ${client.nom}`, mL + 2, y + 5);
      // Horaires à droite
      const arr = d.heureArrivee ? new Date(d.heureArrivee).toLocaleTimeString('fr-FR') : '--';
      const dep = d.heureDepart  ? new Date(d.heureDepart).toLocaleTimeString('fr-FR')  : '--';
      doc.setFontSize(8); doc.setFont(undefined, 'normal'); doc.setTextColor(100,100,100);
      doc.text(`${arr} -> ${dep}`, pageW - mR - 2, y + 5, { align: 'right' });
      doc.setTextColor(0,0,0);
      y += 9;

      // Adresse sur une ligne
      doc.setFontSize(8); doc.setFont(undefined, 'normal'); doc.setTextColor(120,120,120);
      doc.text(client.adresse || '', mL + 2, y); y += 5;
      doc.setTextColor(0,0,0);

      // ── Disposition : tableau à gauche, photos à droite ──
      const totalRamasses = d.cabrisReprisTotal ?? 0;
      const tableW  = 100;          // largeur tableau (mm)
      const photoZoneX = mL + tableW + 4;  // x départ zone photos
      const photoZoneW = contW - tableW - 4; // largeur zone photos (~82mm)
      const photoW  = photoZoneW;            // 1 photo par ligne, pleine largeur zone
      const photoH  = photoW * 4 / 3;       // ratio 3:4 portrait (hauteur > largeur)

      const yDebut = y; // y de départ pour aligner tableau et photos

      // Tableau services (colonne gauche)
      autoTable(doc, {
        startY: yDebut,
        head: [['Service', 'Liv.', 'Ram.']],
        body: d.services.map((s, si) => [s.nom, String(s.cabrisPrevu), si === 0 ? String(totalRamasses) : '']),
        margin: { left: mL },
        tableWidth: tableW,
        headStyles: { fillColor: [37, 99, 235], fontSize: 7, cellPadding: 1.5 },
        bodyStyles: { fontSize: 7, cellPadding: 1.5 },
        columnStyles: {
          0: { cellWidth: tableW * 0.58 },
          1: { cellWidth: tableW * 0.21, halign: 'center' },
          2: { cellWidth: tableW * 0.21, halign: 'center', fontStyle: 'bold', textColor: [146,64,14], fillColor: [254,243,199] },
        },
      });
      const yFinTableau = doc.lastAutoTable.finalY;

      // Commentaire sous le tableau (colonne gauche)
      const texteCommentaire = d.commentaire && d.commentaire.trim() ? d.commentaire.trim() : '';
      let yFinGauche = yFinTableau + 2;
      if (texteCommentaire) {
        const lignes = doc.splitTextToSize(texteCommentaire, tableW - 4);
        const hBox = lignes.length * 4 + 5;
        doc.setFillColor(255, 252, 232);
        doc.rect(mL, yFinGauche, tableW, hBox, 'F');
        doc.setFontSize(6); doc.setFont(undefined, 'bold'); doc.setTextColor(146,64,14);
        doc.text('Commentaire :', mL + 2, yFinGauche + 4);
        doc.setFont(undefined, 'italic'); doc.setTextColor(60,60,60);
        lignes.forEach((l, li) => doc.text(l, mL + 2, yFinGauche + 4 + (li + 1) * 4));
        doc.setFont(undefined, 'normal'); doc.setTextColor(0,0,0);
        yFinGauche += hBox + 2;
      }

      // Photos (colonne droite, alignées avec le haut du tableau)
      let yPhoto = yDebut;
      if (d.photos && d.photos.length > 0) {
        d.photos.forEach((photo, pi) => {
          try {
            const xPos = photoZoneX;  // toujours à gauche, 1 photo par ligne
            if (pi > 0) { yPhoto += photoH + 5; }
            if (yPhoto + photoH > 275) { doc.addPage(); yPhoto = 15; }
            doc.addImage(photo.data, 'JPEG', xPos, yPhoto, photoW, photoH);
            doc.setFontSize(5.5); doc.setTextColor(80,80,80);
            doc.text(photo.heure, xPos, yPhoto + photoH + 3);
            doc.setTextColor(0,0,0);
          } catch(e) { console.warn('Photo PDF:', e); }
        });
        yPhoto += photoH + 5;
      }

      // y final = le plus bas entre fin gauche et fin photos
      y = Math.max(yFinGauche, yPhoto);

      // Séparateur léger entre clients
      doc.setDrawColor(220, 220, 220); doc.setLineWidth(0.3);
      doc.line(mL, y, pageW - mR, y);
      y += 5;
    });

    // ── Total général en bas du PDF ──
    const totalLivresPDF   = tournee.clients.reduce((t, c) => t + c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0), 0);
    const totalRamassesePDF = tournee.clients.reduce((t, c) => {
      const d = clientData[c.id];
      return t + (d ? (d.cabrisReprisTotal ?? 0) : 0);
    }, 0);

    if (y > 250) { doc.addPage(); y = 20; }
    y += 4;
    doc.setDrawColor(37, 99, 235);
    doc.setLineWidth(0.8);
    doc.line(14, y, 196, y); y += 8;

    doc.setFontSize(13); doc.setFont(undefined, 'bold');
    doc.text('TOTAL GENERAL DE LA TOURNEE', 105, y, { align: 'center' }); y += 8;

    autoTable(doc, {
      startY: y,
      head: [['', 'Cabris livres', 'Cabris ramasses']],
      body: [['TOTAL', String(totalLivresPDF), String(totalRamassesePDF)]],
      margin: { left: 14 },
      headStyles: { fillColor: [37, 99, 235] },
      bodyStyles: { fontStyle: 'bold', fillColor: [219, 234, 254], fontSize: 12 },
      styles: { halign: 'center' },
      columnStyles: { 0: { halign: 'left' } }
    });

    return doc;
  };

  const genererPDF = () => {
    const doc = creerDocPDF();
    const maintenant = new Date();
    const dateFichier = maintenant.toISOString().split('T')[0];
    const heureFichier = maintenant.toTimeString().slice(0,8).replace(/:/g,'h');
    doc.save(`Rapport_${dateFichier}_${heureFichier}.pdf`);
  };

  // Générer le PDF et ouvrir le client mail
  const genererEtEnvoyer = () => {
    const emails = [cfg.email1, cfg.email2, cfg.email3].filter(Boolean);
    if (emails.length === 0) {
      alert('Aucun email configuré. Allez dans la Configuration.');
      return;
    }
    // 1. Télécharger le PDF
    genererPDF();
    // 2. Boucler la tournée : effacer la session APRÈS génération du PDF
    // 2. Ouvrir le client mail
    setTimeout(() => {
      const maintenant = new Date();
    const nomFichier = `Rapport_${maintenant.toISOString().split('T')[0]}_${maintenant.toTimeString().slice(0,8).replace(/:/g,'h')}.pdf`;
      const sujet = encodeURIComponent(`Rapport tournée BDL - ${agentName} - ${new Date(tournee.date).toLocaleDateString('fr-FR')}`);
      const corps = encodeURIComponent(
        `Bonjour,

Veuillez trouver ci-joint le rapport de tournée (${nomFichier}).

` +
        `Chauffeur : ${agentName}
` +
        `Date : ${new Date(tournee.date).toLocaleDateString('fr-FR')}
` +
        `Durée : ${Math.floor(dureeTotal/60)}h ${dureeTotal%60}min
` +
        `Clients : ${tournee.clients.length}

Cordialement`
      );
      window.open(`mailto:${emails.join(',')}?subject=${sujet}&body=${corps}`);
      // Session effacée : tournée bouclée
      localStorage.removeItem('bdl-session');
    }, 800);
  };

  return (
    <div style={{ minHeight:'100vh', background:'#f3f4f6', padding:'24px' }}>
      <div style={{ maxWidth:'700px', margin:'0 auto', background:'white', borderRadius:'20px', padding:'36px', boxShadow:'0 4px 20px rgba(0,0,0,0.1)' }}>
        <h2 style={{ fontSize:'28px', fontWeight:'bold', textAlign:'center', marginBottom:'8px' }}>
          <FileText size={32} style={{ verticalAlign:'middle', marginRight:'10px' }} />
          Rapport de Tournée
        </h2>
        {nomTournee && <p style={{ textAlign:'center', color:'#6b7280', fontSize:'16px', marginBottom:'20px' }}>📄 {nomTournee}</p>}

        {/* Horaires récap */}
        <div style={{ background:'#f9fafb', border:'2px solid #e5e7eb', borderRadius:'12px', padding:'16px', marginBottom:'24px' }}>
          <p style={{ fontWeight:'bold', fontSize:'15px', marginBottom:'10px', color:'#374151' }}>⏱️ Chronologie</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'8px', fontSize:'14px', color:'#6b7280' }}>
            {logChargement && <span>📦 Début chargement : <strong>{new Date(logChargement.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {logDepart     && <span>🚚 Départ unité : <strong>{new Date(logDepart.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {logArrivee    && <span>🏠 Arrivée unité : <strong>{new Date(logArrivee.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {logFin        && <span>✅ Fin déchargement : <strong>{new Date(logFin.timestamp).toLocaleTimeString('fr-FR')}</strong></span>}
            {dureeDecharge !== null && <span>🔄 Durée déchargement : <strong>{dureeDecharge} min</strong></span>}
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


        <button onClick={onNouvelle}
          style={{ width:'100%', padding:'22px', background:'#4b5563', color:'white', border:'none', borderRadius:'14px', fontSize:'20px', fontWeight:'bold', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'10px' }}>
          <Home size={24} /> Nouvelle tournee
        </button>

        {/* ── Bilan QR (affiché uniquement si mode QR activé) ── */}
        {(() => {
          // Calcul bilan chargement
          const totalAttenduChgt = tournee.clients.reduce((t, c) =>
            t + c.services.reduce((s, sv) => s + (sv.cabrisIds?.length || 0), 0), 0);
          const modeQR = totalAttenduChgt > 0;
          if (!modeQR) return null;

          const totalScannésChgt  = cabrisChargement?.scannés?.length || 0;
          const anomaliesChgt     = cabrisChargement?.anomalies || [];

          // Calcul bilan livraison (agrégation sur tous les clients)
          const totalAttenduLiv   = totalAttenduChgt; // même total
          const totalScannésLiv   = tournee.clients.reduce((t, c) => {
            const d = clientData[c.id];
            if (!d || !d.cabrisScannés) return t;
            return t + Object.values(d.cabrisScannés).reduce((s, arr) => s + arr.length, 0);
          }, 0);
          const anomaliesLivAll   = tournee.clients.flatMap(c => clientData[c.id]?.anomaliesLivraison || []);

          const chgtOK   = totalScannésChgt >= totalAttenduChgt && anomaliesChgt.length === 0;
          const livOK    = totalScannésLiv  >= totalAttenduLiv  && anomaliesLivAll.length === 0;

          return (
            <div style={{ background:'#f0f9ff', border:'2px solid #7dd3fc', borderRadius:'14px', padding:'20px', marginBottom:'24px' }}>
              <p style={{ fontSize:'17px', fontWeight:'bold', marginBottom:'14px', color:'#0369a1' }}>
                📱 Bilan des scans QR
              </p>
              <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px', marginBottom:'14px' }}>
                {/* Chargement */}
                <div style={{ background:'white', borderRadius:'10px', padding:'14px', textAlign:'center', border:`2px solid ${chgtOK ? '#86efac' : '#fca5a5'}` }}>
                  <p style={{ fontSize:'12px', color:'#6b7280', margin:'0 0 4px' }}>Chargement</p>
                  <p style={{ fontSize:'22px', fontWeight:'bold', margin:'0 0 4px', color: chgtOK ? '#16a34a' : '#dc2626' }}>
                    {totalScannésChgt}/{totalAttenduChgt} {chgtOK ? '✅' : '⚠️'}
                  </p>
                  {anomaliesChgt.length > 0 && (
                    <p style={{ fontSize:'11px', color:'#dc2626', margin:0 }}>{anomaliesChgt.length} anomalie(s)</p>
                  )}
                </div>
                {/* Livraison */}
                <div style={{ background:'white', borderRadius:'10px', padding:'14px', textAlign:'center', border:`2px solid ${livOK ? '#86efac' : '#fca5a5'}` }}>
                  <p style={{ fontSize:'12px', color:'#6b7280', margin:'0 0 4px' }}>Livraison</p>
                  <p style={{ fontSize:'22px', fontWeight:'bold', margin:'0 0 4px', color: livOK ? '#16a34a' : '#dc2626' }}>
                    {totalScannésLiv}/{totalAttenduLiv} {livOK ? '✅' : '⚠️'}
                  </p>
                  {anomaliesLivAll.length > 0 && (
                    <p style={{ fontSize:'11px', color:'#dc2626', margin:0 }}>{anomaliesLivAll.length} anomalie(s)</p>
                  )}
                </div>
              </div>
              {/* Détail anomalies chargement */}
              {anomaliesChgt.length > 0 && (
                <div style={{ marginBottom:'10px' }}>
                  <p style={{ fontSize:'13px', fontWeight:'600', color:'#dc2626', margin:'0 0 6px' }}>⚠️ Anomalies chargement :</p>
                  {anomaliesChgt.map((a, i) => (
                    <div key={i} style={{ background:'#fef2f2', border:'1px solid #fca5a5', borderRadius:'8px', padding:'8px 12px', marginBottom:'4px', fontSize:'13px' }}>
                      <strong>{a.id}</strong> — {a.commentaire}
                    </div>
                  ))}
                </div>
              )}
              {/* Détail anomalies livraison */}
              {anomaliesLivAll.length > 0 && (
                <div>
                  <p style={{ fontSize:'13px', fontWeight:'600', color:'#dc2626', margin:'0 0 6px' }}>⚠️ Anomalies livraison :</p>
                  {anomaliesLivAll.map((a, i) => (
                    <div key={i} style={{ background:'#fef2f2', border:'1px solid #fca5a5', borderRadius:'8px', padding:'8px 12px', marginBottom:'4px', fontSize:'13px' }}>
                      <strong>{a.id}</strong> — {a.commentaire}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })()}

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
                {/* Services livres - lecture seule */}
                <div style={{ marginBottom:'8px' }}>
                  {d.services.map(s => (
                    <div key={s.id} style={{ display:'flex', justifyContent:'space-between', padding:'5px 10px', background:'white', borderRadius:'6px', marginBottom:'4px', fontSize:'14px' }}>
                      <span style={{ color:'#374151' }}>{s.nom}</span>
                      <span style={{ color:'#2563eb', fontWeight:'600' }}>{s.cabrisPrevu} livres</span>
                    </div>
                  ))}
                </div>
                {/* Total ramasses - une seule valeur */}
                <div style={{ background:'#fef3c7', border:'2px solid #fde68a', borderRadius:'8px', padding:'8px 14px', marginBottom:'8px', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                  <span style={{ fontWeight:'bold', color:'#92400e', fontSize:'14px' }}>Total ramasses :</span>
                  <span style={{ fontWeight:'bold', color:'#92400e', fontSize:'20px' }}>{d.cabrisReprisTotal ?? 0}</span>
                </div>

                <div style={{ background:'#fefce8', border:'1px solid #fde68a', borderRadius:'8px', padding:'10px 12px', marginBottom:'8px' }}>
                  <p style={{ fontSize:'12px', fontWeight:'600', color:'#92400e', margin:'0 0 4px' }}>Commentaire</p>
                  <p style={{ fontSize:'13px', color:'#374151', margin:0, fontStyle:'italic' }}>
                    {d.commentaire && d.commentaire.trim() ? d.commentaire : <span style={{color:'#9ca3af'}}>Aucun commentaire</span>}
                  </p>
                </div>

                {d.photos && d.photos.length > 0 && (
                  <div>
                    <p style={{ fontSize:'13px', fontWeight:'600', color:'#374151', margin:'0 0 6px' }}>{d.photos.length} photo(s)</p>
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

          {/* Total général tous clients */}
          {(() => {
            const totalLivres   = tournee.clients.reduce((t, c) => t + c.services.reduce((s, sv) => s + sv.cabrisPrevu, 0), 0);
            const totalRamasses = tournee.clients.reduce((t, c) => {
              const d = clientData[c.id];
              return t + (d ? (d.cabrisReprisTotal ?? 0) : 0);
            }, 0);
            return (
              <div style={{ background:'linear-gradient(135deg,#1d4ed8,#1e40af)', borderRadius:'14px', padding:'20px 24px', marginTop:'8px' }}>
                <p style={{ color:'white', fontWeight:'bold', fontSize:'16px', margin:'0 0 14px', textAlign:'center' }}>
                  TOTAL GENERAL DE LA TOURNEE
                </p>
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px' }}>
                  <div style={{ background:'rgba(255,255,255,0.15)', borderRadius:'10px', padding:'14px', textAlign:'center' }}>
                    <p style={{ color:'#bfdbfe', fontSize:'12px', margin:'0 0 4px' }}>Cabris livres</p>
                    <p style={{ color:'white', fontSize:'32px', fontWeight:'bold', margin:0 }}>{totalLivres}</p>
                  </div>
                  <div style={{ background:'rgba(255,255,255,0.15)', borderRadius:'10px', padding:'14px', textAlign:'center' }}>
                    <p style={{ color:'#bfdbfe', fontSize:'12px', margin:'0 0 4px' }}>Cabris ramasses</p>
                    <p style={{ color:'white', fontSize:'32px', fontWeight:'bold', margin:0 }}>{totalRamasses}</p>
                  </div>
                </div>
                <p style={{ color:'#bfdbfe', fontSize:'12px', textAlign:'center', margin:'12px 0 0' }}>
                  {totalRamasses} cabris sales dans le camion
                </p>
              </div>
            );
          })()}
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
  const [nomTournee, setNomTournee]     = useState('');
  const [cabrisChargement, setCabrisChargement] = useState(null); // { scannés: [], anomalies: [] }

  // Chargement de la config au démarrage (depuis config.json + localStorage)
  useEffect(() => {
    chargerConfigAsync().then(cfg => setConfig(cfg));
  }, []);

  // ── Restauration automatique au démarrage (plantage/fermeture app) ──
  useEffect(() => {
    try {
      const raw = localStorage.getItem('bdl-session');
      if (!raw) return;
      const s = JSON.parse(raw);
      // Restaurer seulement si une tournée active était en cours
      if (s && s.tournee && !['login','import','config','rapport'].includes(s.vue)) {
        setVue(s.vue);
        setAgentName(s.agentName || '');
        setTournee(s.tournee);
        setNomTournee(s.nomTournee || '');
        setClientIndex(s.clientIndex || 0);
        setStartTime(s.startTime || null);
        setLogs(s.logs || []);
        setClientData(s.clientData || {});
        setCabrisChargement(s.cabrisChargement || null);
      }
    } catch(e) {
      console.warn('Erreur restauration session:', e);
    }
  }, []); // [] = exécuté une seule fois au montage

  // ── Sauvegarde automatique à chaque changement d'état ──
  useEffect(() => {
    if (!['login','import','config'].includes(vue) && tournee) { // 'rapport' inclus intentionnellement
      try {
        // On sauvegarde sans les photos (trop lourdes) pour éviter le dépassement localStorage
        const clientDataSansPhotos = Object.fromEntries(
          Object.entries(clientData).map(([k, v]) => [k, { ...v, photos: [] }])
        );
        const session = { vue, agentName, tournee, nomTournee, clientIndex, startTime, logs, clientData: clientDataSansPhotos, cabrisChargement };
        localStorage.setItem('bdl-session', JSON.stringify(session));
      } catch(e) {
        console.warn('Sauvegarde session impossible:', e);
      }
    }
    // Effacer la session uniquement quand on revient au login (nouvelle tournée)
    if (vue === 'login') {
      localStorage.removeItem('bdl-session');
    }
  }, [vue, agentName, tournee, nomTournee, clientIndex, startTime, logs, clientData, cabrisChargement]);

  const addLog = (type, idx = null, extra = {}) => {
    const entry = { type, clientIndex: idx, timestamp: new Date().toISOString(), ...extra };
    setLogs(prev => [...prev, entry]);
    return entry;
  };

  const sessionEnCours = () => {
    try {
      const raw = localStorage.getItem('bdl-session');
      if (!raw) return null;
      const s = JSON.parse(raw);
      return s.tournee ? s : null;
    } catch { return null; }
  };

  const restaurerSession = (s) => {
    setVue(s.vue); setAgentName(s.agentName || '');
    setTournee(s.tournee); setNomTournee(s.nomTournee || '');
    setClientIndex(s.clientIndex || 0); setStartTime(s.startTime || null);
    setLogs(s.logs || []); setClientData(s.clientData || {});
    setCabrisChargement(s.cabrisChargement || null);
  };

  const handleLogin = (nom, reprise = false) => {
    if (reprise) {
      const s = sessionEnCours();
      if (s) { restaurerSession(s); return; }
    }
    // Vérifier si une session non bouclée existe
    const s = sessionEnCours();
    if (s) {
      const nomTour = s.nomTournee || 'en cours';
      const ok = window.confirm(
        `La tournee "${nomTour}" n'a pas ete bouclee.

OK      = Reprendre la tournee
Annuler = Demarrer une nouvelle tournee`
      );
      if (ok) { restaurerSession(s); return; }
      // Abandonner : effacer la session
      localStorage.removeItem('bdl-session');
    }
    setAgentName(nom); setVue('import');
  };
  const handleConfig   = ()    => setVue('config');
  const handleImport = (t, nom) => {
    const s = sessionEnCours();
    if (s) {
      const nomTour = s.nomTournee || 'en cours';
      const ok = window.confirm(
        `La tournee "${nomTour}" n'a pas ete bouclee.

OK      = Reprendre la tournee
Annuler = Charger la nouvelle tournee`
      );
      if (ok) { restaurerSession(s); return; }
      localStorage.removeItem('bdl-session');
    }
    setTournee(t); setNomTournee(nom || ''); setClientIndex(0); setClientData({}); setLogs([]); setCabrisChargement(null); setVue('recap');
  };

  const handleDemarrer = (clientsOrdres, heureChargement, cabrisChargData) => {
    setTournee(prev => ({ ...prev, clients: clientsOrdres }));
    const now = new Date().toISOString();
    setStartTime(heureChargement); // l'horodatage démarre dès le chargement
    setCabrisChargement(cabrisChargData || { scannés: [], anomalies: [] });
    setLogs([
      { type: 'DEBUT_CHARGEMENT', clientIndex: null, timestamp: heureChargement },
      { type: 'DEPART_UNITE',     clientIndex: null, timestamp: now }
    ]);
    setVue('tournee');
  };

  const handleArrivee = () => { addLog('ARRIVEE_CLIENT', clientIndex); setVue('client'); };

  const handleDepartClient = (commentaire, services, photos, cabrisReprisTotal, qrLivraison) => {
    const now = new Date().toISOString();
    const client = tournee.clients[clientIndex];
    const heureArrivee = [...logs].reverse().find(l => l.type === 'ARRIVEE_CLIENT' && l.clientIndex === clientIndex)?.timestamp;
    setClientData(prev => ({
      ...prev,
      [client.id]: {
        services,
        commentaire,
        photos: photos || [],
        heureArrivee,
        heureDepart: now,
        cabrisReprisTotal: cabrisReprisTotal || 0,
        cabrisScannés: qrLivraison?.cabrisScannés || {},
        anomaliesLivraison: qrLivraison?.anomalies || []
      }
    }));
    addLog('DEPART_CLIENT', clientIndex, { commentaire });
    if (clientIndex < tournee.clients.length - 1) {
      setClientIndex(clientIndex + 1); setVue('tournee');
    } else {
      setVue('retour');
    }
  };

  const handleArriveeUnite   = () => { addLog('ARRIVEE_UNITE');  setVue('decharge'); };
  const handleDecharge       = () => { addLog('FIN_DECHARGE');   setVue('rapport');  };
  const handleNouvelle = () => {
    localStorage.removeItem('bdl'); localStorage.removeItem('bdl-session');
    setVue('login'); setTournee(null); setNomTournee(''); setAgentName('');
    setLogs([]); setClientData({}); setStartTime(null); setClientIndex(0); setCabrisChargement(null);
  };

  const handleNouvelleAvecCheck = () => {
    // Proposer reprise si tournée non bouclée (rapport non généré)
    const s = sessionEnCours();
    if (s) {
      const nom = s.nomTournee || 'en cours';
      const reprendre = window.confirm(
        `La tournee "${nom}" n'a pas ete bouclee (rapport non genere).

OK      = Reprendre la tournee
Annuler = Abandonner et demarrer une nouvelle tournee`
      );
      if (reprendre) { restaurerSession(s); return; }
    }
    handleNouvelle();
  };

  // config est chargé depuis config.json + localStorage via le state

  if (vue === 'config')  return <VueConfig config={config} onSave={(cfg) => { setConfig(cfg); }} onRetour={() => setVue(agentName ? 'import' : 'login')} />;
  if (vue === 'login')   return <VueLogin  onLogin={handleLogin} onConfig={handleConfig} />;
  if (vue === 'import')  return <VueImport key={Date.now()} onImport={handleImport} onConfig={handleConfig} />;
  if (vue === 'recap')   return <VueRecap  tournee={tournee} onDemarrer={handleDemarrer} onRetour={() => setVue('import')} />;
  if (vue === 'retour')  return <VueRetourUnite onArriveeUnite={handleArriveeUnite} adresseUnite={config.adresseUnite} logs={logs} />;
  if (vue === 'decharge') return <VueDecharge   onDechargeTerminee={handleDecharge} logs={logs} />;
  if (vue === 'rapport') return <VueRapport tournee={tournee} clientData={clientData} logs={logs} startTime={startTime} agentName={agentName} onNouvelle={handleNouvelleAvecCheck} nomTournee={nomTournee} cabrisChargement={cabrisChargement} />;

  if (vue === 'tournee' && tournee)
    return <VueTournee client={tournee.clients[clientIndex]} index={clientIndex} total={tournee.clients.length} onArrivee={handleArrivee} />;
  if (vue === 'client' && tournee)
    return <VueClient key={`client-${clientIndex}`} client={tournee.clients[clientIndex]} onDepart={handleDepartClient} tournee={tournee} />;

  return null;
}
