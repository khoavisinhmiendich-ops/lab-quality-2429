import { initializeApp, getApps, getApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyCf7ZWWWf5NaP6HSFvppKRAJV5cOteI9WY",
  authDomain: "qlcl-vi-sinh---mien-dich.firebaseapp.com",
  projectId: "qlcl-vi-sinh---mien-dich",
  storageBucket: "qlcl-vi-sinh---mien-dich.firebasestorage.app",
  messagingSenderId: "1028392607937",
  appId: "1:1028392607937:web:eac8992667bb51d777f440",
  measurementId: "G-1EJX1GD8BF"
};

const app = getApps().length > 0 ? getApp() : initializeApp(firebaseConfig);
export const db = getFirestore(app);