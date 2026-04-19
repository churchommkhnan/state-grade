import { cn } from "@/lib/utils";

export function Card({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn(
        "rounded-2xl border border-white/15 bg-slate-900/60 backdrop-blur-xl shadow-[0_10px_30px_-12px_rgba(99,102,241,0.45)]",
        className
      )}
      {...props}
    />
  );
}
