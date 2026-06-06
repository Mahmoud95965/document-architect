import { motion } from "framer-motion";

const STEPS = [
  "Architecting your document",
  "Drafting sections",
  "Composing typography",
  "Finalizing layout",
];

export function LoadingState() {
  return (
    <div className="flex flex-col items-center justify-center gap-6 py-12">
      <div className="relative h-20 w-20">
        <motion.div
          className="absolute inset-0 rounded-full border-2 border-transparent"
          style={{
            borderTopColor: "oklch(0.7 0.22 265)",
            borderRightColor: "oklch(0.7 0.22 320)",
          }}
          animate={{ rotate: 360 }}
          transition={{ duration: 1.6, repeat: Infinity, ease: "linear" }}
        />
        <motion.div
          className="absolute inset-2 rounded-full border border-white/10"
          animate={{ rotate: -360 }}
          transition={{ duration: 2.4, repeat: Infinity, ease: "linear" }}
        />
        <div className="absolute inset-0 m-auto h-2 w-2 rounded-full bg-white/80" />
      </div>
      <div className="h-6 overflow-hidden text-center">
        <motion.div
          animate={{ y: [0, -24, -48, -72, 0] }}
          transition={{ duration: 6, repeat: Infinity, ease: "easeInOut" }}
        >
          {STEPS.map((s) => (
            <div key={s} className="h-6 text-sm text-muted-foreground">
              {s}…
            </div>
          ))}
        </motion.div>
      </div>
    </div>
  );
}
