import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
const e = import.meta.env;
console.log(e.VITE_FB_API_KEY);
const app = initializeApp({
  apiKey: e.VITE_FB_API_KEY,
  authDomain: e.VITE_FB_AUTH_DOMAIN,
  projectId: e.VITE_FB_PROJECT_ID,
  appId: e.VITE_FB_APP_ID,
});
export const auth = getAuth(app);
export const db = getFirestore(app);
export const ADMINS = (e.VITE_ADMIN_EMAILS || "")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);
