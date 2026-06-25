"use client";
import { Check } from "lucide-react";

interface Step {
  id: number;
  label: string;
}

interface Props {
  steps: Step[];
  current: number;
}

export default function StepIndicator({ steps, current }: Props) {
  return (
    <div className="flex items-center justify-center gap-0 flex-wrap">
      {steps.map((step, idx) => {
        const done = step.id < current;
        const active = step.id === current;
        return (
          <div key={step.id} className="flex items-center">
            <div className="flex flex-col items-center gap-1">
              <div
                className={`
                  w-9 h-9 rounded-full flex items-center justify-center
                  font-bold text-sm transition-all duration-300
                  ${done ? "text-white" : ""}
                  ${active ? "text-white ring-4 ring-pink-100" : ""}
                  ${!done && !active ? "bg-slate-200 text-slate-500" : ""}
                `}
                style={
                  done
                    ? { background: "#72BF44" }
                    : active
                    ? { background: "linear-gradient(135deg, #E91E8C, #F15A22)" }
                    : {}
                }
              >
                {done ? <Check className="w-4 h-4" /> : step.id}
              </div>
              <span
                className={`text-xs font-medium whitespace-nowrap ${
                  active ? "text-[#E91E8C]" : done ? "text-[#72BF44]" : "text-slate-400"
                }`}
              >
                {step.label}
              </span>
            </div>
            {idx < steps.length - 1 && (
              <div
                className={`w-12 h-0.5 mb-4 mx-1 transition-colors duration-300 ${
                  done ? "bg-[#72BF44]" : "bg-slate-200"
                }`}
              />
            )}
          </div>
        );
      })}
    </div>
  );
}
