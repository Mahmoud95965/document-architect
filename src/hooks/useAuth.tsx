import {
  createContext,
  useContext,
  useEffect,
  useState,
  type ReactNode,
} from "react";
import {
  onAuthStateChanged,
  signOut as firebaseSignOut,
  type User,
} from "firebase/auth";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { auth, db } from "@/lib/firebase";

// ─── Types ─────────────────────────────────────────────────
export type Plan = "pro" | "free" | null;

interface AuthContextType {
  user:               User | null;
  plan:               Plan;
  filesCountToday:    number;
  lastGenerationDate: string;
  loading:            boolean;
  signOut:            () => Promise<void>;
  incrementDailyFiles: () => Promise<number>;
}

// ─── Context ────────────────────────────────────────────────
const AuthContext = createContext<AuthContextType>({
  user:               null,
  plan:               null,
  filesCountToday:    0,
  lastGenerationDate: "",
  loading:            true,
  signOut:            async () => {},
  incrementDailyFiles: async () => 0,
});

// ─── Provider ───────────────────────────────────────────────
export function AuthProvider({ children }: { children: ReactNode }) {
  const [user,               setUser]               = useState<User | null>(null);
  const [plan,               setPlan]               = useState<Plan>(null);
  const [filesCountToday,    setFilesCountToday]    = useState<number>(0);
  const [lastGenerationDate, setLastGenerationDate] = useState<string>("");
  const [loading,            setLoading]            = useState(true);

  // Helper to get today's date in local YYYY-MM-DD format
  const getTodayStr = () => {
    const d = new Date();
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  };

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (firebaseUser) => {
      setUser(firebaseUser);

      if (firebaseUser) {
        try {
          console.log("[useAuth] Fetching plan from Firestore for UID:", firebaseUser.uid);
          const snap = await getDoc(doc(db, "users", firebaseUser.uid));
          if (snap.exists()) {
            const data = snap.data();
            const rawPlan = data?.plan as string | undefined;
            console.log("[useAuth] Firestore Document found! Data:", data, "rawPlan:", rawPlan);
            setPlan(rawPlan === "pro" ? "pro" : "free");

            // Extract daily limit data
            const todayStr = getTodayStr();
            const lastDate = data?.lastGenerationDate as string | undefined;
            const count = data?.filesCountToday as number | undefined;

            if (lastDate !== todayStr) {
              // It's a new day, count resets to 0 under the hood
              setFilesCountToday(0);
              setLastGenerationDate(todayStr);
            } else {
              setFilesCountToday(count ?? 0);
              setLastGenerationDate(lastDate ?? todayStr);
            }
          } else {
            console.log("[useAuth] No document in 'users' collection for UID:", firebaseUser.uid, "→ Defaulting to 'free'");
            setPlan("free");
            setFilesCountToday(0);
            setLastGenerationDate(getTodayStr());
          }
        } catch (err) {
          console.error("[useAuth] Error fetching document from Firestore:", err);
          setPlan("free");
          setFilesCountToday(0);
          setLastGenerationDate(getTodayStr());
        }
      } else {
        setPlan(null);
        setFilesCountToday(0);
        setLastGenerationDate("");
      }

      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  const signOut = async () => {
    await firebaseSignOut(auth);
    setUser(null);
    setPlan(null);
    setFilesCountToday(0);
    setLastGenerationDate("");
  };

  const incrementDailyFiles = async (): Promise<number> => {
    if (!user) return 0;
    const todayStr = getTodayStr();
    let newCount = 1;

    if (lastGenerationDate === todayStr) {
      newCount = filesCountToday + 1;
    }

    try {
      await setDoc(
        doc(db, "users", user.uid),
        {
          filesCountToday: newCount,
          lastGenerationDate: todayStr,
        },
        { merge: true }
      );
      setFilesCountToday(newCount);
      setLastGenerationDate(todayStr);
      return newCount;
    } catch (err) {
      console.error("[useAuth] Error writing filesCount to Firestore:", err);
      // Fallback update to local state even if Firestore write fails
      setFilesCountToday(newCount);
      setLastGenerationDate(todayStr);
      return newCount;
    }
  };

  return (
    <AuthContext.Provider
      value={{
        user,
        plan,
        filesCountToday,
        lastGenerationDate,
        loading,
        signOut,
        incrementDailyFiles,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
}

// ─── Hook ───────────────────────────────────────────────────
export const useAuth = () => useContext(AuthContext);
