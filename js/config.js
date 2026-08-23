/**
 * Configuración global de la aplicación
 * 
 * Si despliegas en Netlify, puedes configurar la variable de entorno:
 * GOOGLE_CLIENT_ID = tu-client-id.apps.googleusercontent.com
 * 
 * Para desarrollo local o despliegue estático directo, puedes colocar tu Client ID
 * aquí en DEFAULT_GOOGLE_CLIENT_ID.
 */

window.APP_CONFIG = {
  // Puedes pegar aquí tu Client ID para desarrollo local:
  DEFAULT_GOOGLE_CLIENT_ID: '',

  // Endpoint de función serverless en Netlify (opcional si usas variables de entorno)
  NETLIFY_CONFIG_ENDPOINT: '/.netlify/functions/config'
};
