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
  // Pega aquí tu Client ID de Google Cloud Console:
  // Ejemplo: '123456789-abcde.apps.googleusercontent.com'
  DEFAULT_GOOGLE_CLIENT_ID: 'TU_CLIENT_ID_AQUI',

  // Clave Interna de Google Gemini API (Tier 100% Gratuito):
  // Al estar aquí configurada por interno, ningún usuario ni gerente tendrá que configurar nada en pantalla.
  GEMINI_API_KEY: 'AQ.Ab8RN6K2Uj-kWD8TILayQ0hP2bUsqtzAlk3rUhgAzGGFmx5DLA',
};
