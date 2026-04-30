const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

// Text replacements: template text → field key from request body
const REPLACEMENTS = [
  { from: 'Creación de Contenido (Video & Foto)+ Brand assesement ', key: 'nombre_servicio' },
  { from: 'Creación de Contenido (Video & Foto)+ Brand assesement', key: 'nombre_servicio' },
  { from: 'Creación de contenido audiovisual en formato vertical, optimizado para redes sociales, alineado con la estrategia de la marca y pensado para facilitar su ejecución y publicación.', key: 'descripcion' },
  { from: 'Acompañamiento estratégico para ayudar a una marca a identificar su posición real en el mercado, clarificar su voz y diseñar una hoja de ruta de contenido accionable. A través de un proceso estructurado de reuniones y feedback, definimos dónde está la marca hoy, dónde debería estar, y cómo llegar ahí', key: 'incluye' },
  { from: 'Plazo: 6 semanas', key: 'plazo', prefix: 'Plazo: ' },
  { from: '6 semanas', key: 'plazo' },
  { from: "Brand Discovery Document: análisis de competidores, oportunidad en el mercado digital, tono de voz, do's &amp; don'ts", key: 'entregable_1' },
  { from: 'Content Guide (DIY): guía de contenido que el propio equipo de la marca puede producir', key: 'entregable_2' },
  { from: 'Filmación y edición de 6 videos (Reels) con contenido previamente definido, ', key: 'entregable_3' },
  { from: 'Filmación y edición de 6 videos (Reels) con contenido previamente definido,', key: 'entregable_3' },
  { from: '60 fotografías de producto y de situación, pensadas para feed, stories y otros usos digitales.', key: 'entregable_4' },
  { from: 'Dirección de arte, asegurando coherencia visual, estética de marca y calidad en todo el contenido.', key: 'entregable_5' },
  { from: 'CON ESTE PACK SE PREVEE CUBRIR EL CONTENIDO PARA UN MÍNIMO DE 3 MESES.', key: 'nota_aclaracion' },
  { from: '2.500€', key: 'precio' },
  { from: '2.500&#x20AC;', key: 'precio' },
  { from: 'Brand discovery', key: 'timeline_1' },
  { from: 'Producción contenido', key: 'timeline_2' },
  { from: 'Entrega contenidos', key: 'timeline_3' },
];

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replaceInXml(xml, from, to) {
  // Direct string replacement (XML-safe)
  let result = xml;
  // Try literal replacement first
  if (result.includes(from)) {
    result = result.split(from).join(to);
  }
  // Try with XML-escaped version
  const escapedFrom = escapeXml(from);
  if (escapedFrom !== from && result.includes(escapedFrom)) {
    result = result.split(escapedFrom).join(escapeXml(to));
  }
  return result;
}

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const data = req.body;

    // Load template
    const templatePath = path.join(__dirname, 'template.pptx');
    const templateBuffer = fs.readFileSync(templatePath);

    // Open as ZIP
    const zip = await JSZip.loadAsync(templateBuffer);

    // Process all slide XML files
    const slideFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );

    for (const slideFile of slideFiles) {
      let xml = await zip.files[slideFile].async('string');

      // Apply all replacements
      for (const { from, key, prefix } of REPLACEMENTS) {
        const val = data[key];
        if (!val) continue;
        const to = prefix ? prefix + val : val;
        xml = replaceInXml(xml, from, to);
      }

      zip.file(slideFile, xml);
    }

    // Generate output
    const outputBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    const filename = `Bright_Propuesta_${(data.nombre_servicio || 'Servicio').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30)}.pptx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', outputBuffer.length);

    return res.status(200).send(outputBuffer);

  } catch (err) {
    console.error('Generate error:', err);
    return res.status(500).json({ error: err.message });
  }
};
