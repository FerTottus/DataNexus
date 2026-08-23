/**
 * Netlify Serverless Function: config.js
 * Expone de forma segura las variables públicas de configuración (como GOOGLE_CLIENT_ID)
 * definidas en el panel de Netlify (Site Settings > Environment Variables).
 */

exports.handler = async function (event, context) {
  // Solo permitir solicitudes GET
  if (event.httpMethod !== 'GET') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Método no permitido' })
    };
  }

  const clientId = process.env.GOOGLE_CLIENT_ID || '';

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/json',
      'Access-Control-Allow-Origin': '*'
    },
    body: JSON.stringify({
      clientId: clientId
    })
  };
};
