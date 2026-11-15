// Configuration de l'accès OneDrive.
// Remplir clientId et tenantId avec les valeurs de l'application déclarée dans Microsoft Entra
// (voir GUIDE-MISE-EN-LIGNE.md). Tant que clientId est vide, l'appli fonctionne en « mode local »
// (données uniquement dans ce navigateur, sans OneDrive).
window.CAVE_CONFIG = {
  clientId: '',
  tenantId: '',
  // Emplacement du fichier de données dans votre OneDrive
  cheminFichier: 'Cave/cave.json'
};
