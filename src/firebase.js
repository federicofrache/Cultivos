import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

// 1) Andá a https://console.firebase.google.com -> tu proyecto -> ícono de engranaje
//    -> "Configuración del proyecto" -> pestaña "General" -> sección "Tus apps"
//    -> creá una "App web" (</>) si no tenés una, y copiá el objeto firebaseConfig de ahí.
// 2) Reemplazá los valores de abajo por los tuyos.
const firebaseConfig = {
  apiKey: "AIzaSyD5ThZCKkeFGORCEot1P4YulLzluWedkX4",
  authDomain: "margenes-cultivos.firebaseapp.com",
  projectId: "margenes-cultivos",
  storageBucket: "margenes-cultivos.firebasestorage.app",
  messagingSenderId: "720906475603",
  appId: "1:720906475603:web:a160edb66f1931f1b14ea6",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);
export const storage = getStorage(app);
