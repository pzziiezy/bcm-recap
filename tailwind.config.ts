import type { Config } from "tailwindcss";

export default {
  content: [
    "./app/**/*.{ts,tsx}",
    "./components/**/*.{ts,tsx}",
    "./lib/**/*.{ts,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          50:  "#fdf2f8",
          100: "#fce7f3",
          200: "#fbcfe8",
          500: "#E91E8C",
          600: "#d41679",
          700: "#be185d",
          900: "#831843",
        },
        mini: {
          pink:      "#E91E8C",
          blue:      "#00A6E2",
          yellow:    "#FFD100",
          orange:    "#F15A22",
          green:     "#72BF44",
          darkGreen: "#1A7C3A",
        },
      },
    },
  },
  plugins: [],
} satisfies Config;
