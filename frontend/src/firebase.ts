import { initializeApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut, type User } from 'firebase/auth';
import { addDoc, collection, deleteDoc, doc, getDocs, getFirestore, onSnapshot } from 'firebase/firestore';

const config = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};

export const firebaseEnabled = Object.values(config).every(Boolean);
export const adminEmail = import.meta.env.VITE_FIREBASE_ADMIN_EMAIL || '';
const app = firebaseEnabled ? initializeApp(config) : null;
const auth = app ? getAuth(app) : null;
const store = app ? getFirestore(app) : null;
const RULES_KEY = 'taxwise-name-rules';

export type NameRule = { id: string; source: string; target: string };

export function watchAuth(onChange: (user: User | null) => void) {
  return auth ? onAuthStateChanged(auth, onChange) : () => {};
}

export async function signInOwner(email: string, password: string) {
  if (!auth) throw new Error('ยังไม่ได้ตั้งค่า Firebase');
  return signInWithEmailAndPassword(auth, email, password);
}

export async function signOutOwner() {
  if (auth) await signOut(auth);
}

function localRules(): NameRule[] {
  try { return JSON.parse(localStorage.getItem(RULES_KEY) || '[]') as NameRule[]; }
  catch { return []; }
}

export async function getNameRules(): Promise<NameRule[]> {
  if (!store) return localRules();
  const snapshot = await getDocs(collection(store, 'nameRules'));
  return snapshot.docs.map((item) => ({ id: item.id, source: item.data().source, target: item.data().target }))
    .sort((a, b) => a.source.localeCompare(b.source, 'th'));
}

export function watchNameRules(onChange: (rules: NameRule[]) => void, onError: (error: Error) => void) {
  if (!store) {
    const refresh = () => onChange(localRules());
    window.addEventListener('taxwise-rules-changed', refresh);
    refresh();
    return () => window.removeEventListener('taxwise-rules-changed', refresh);
  }
  let unsubscribe = () => {};
  let cancelled = false;
  if (!cancelled) unsubscribe = onSnapshot(collection(store, 'nameRules'), (snapshot) => {
    onChange(snapshot.docs.map((item) => ({ id: item.id, source: item.data().source, target: item.data().target }))
      .sort((a, b) => a.source.localeCompare(b.source, 'th')));
  }, onError);
  return () => { cancelled = true; unsubscribe(); };
}

export async function createNameRule(source: string, target: string) {
  if (!store) {
    const rules = localRules();
    if (rules.some((rule) => rule.source.toLocaleLowerCase() === source.toLocaleLowerCase())) throw new Error('มีกฎสำหรับชื่อนี้แล้ว');
    rules.push({ id: crypto.randomUUID(), source, target });
    localStorage.setItem(RULES_KEY, JSON.stringify(rules));
    window.dispatchEvent(new Event('taxwise-rules-changed'));
    return;
  }
  if (!auth?.currentUser || auth.currentUser.email?.toLowerCase() !== adminEmail.toLowerCase()) {
    throw new Error('กรุณาเข้าสู่ระบบด้วยบัญชีผู้ดูแลก่อนจัดการกฎ');
  }
  const rules = await getNameRules();
  if (rules.some((rule) => rule.source.trim().toLocaleLowerCase() === source.trim().toLocaleLowerCase())) throw new Error('มีกฎสำหรับชื่อนี้แล้ว');
  await addDoc(collection(store, 'nameRules'), { source, target });
}

export async function deleteNameRule(id: string) {
  if (!store) {
    localStorage.setItem(RULES_KEY, JSON.stringify(localRules().filter((rule) => rule.id !== id)));
    window.dispatchEvent(new Event('taxwise-rules-changed'));
    return;
  }
  if (!auth?.currentUser || auth.currentUser.email?.toLowerCase() !== adminEmail.toLowerCase()) {
    throw new Error('กรุณาเข้าสู่ระบบด้วยบัญชีผู้ดูแลก่อนจัดการกฎ');
  }
  await deleteDoc(doc(store, 'nameRules', id));
}
