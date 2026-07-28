import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

// 1) Andá a https://console.firebase.google.com -> tu proyecto -> ícono de engranaje
//    -> "Configuración del proyecto" -> pestaña "General" -> sección "Tus apps"
//    -> creá una "App web" (</>) si no tenés una, y copiá el objeto firebaseConfig de ahí.
// 2) Reemplazá los valores de abajo por los tuyos.
const firebaseConfig = {
  apiKey: "TU_API_KEY",
  authDomain: "TU_PROYECTO.firebaseapp.com",
  projectId: "TU_PROYECTO",
  storageBucket: "TU_PROYECTO.appspot.com",
  messagingSenderId: "TU_SENDER_ID",
  appId: "TU_APP_ID",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
export const storage = getStorage(app);
