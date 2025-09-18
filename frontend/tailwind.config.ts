import type { Config } from "tailwindcss";

export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["Inter", "system-ui", "Segoe UI", "Roboto", "Helvetica Neue", "Arial", "sans-serif"]
      },
      colors: {
        brand: {
          50: "#effef7",
          100: "#dafeef",
          200: "#b6f9de",
          300: "#83f0c7",
          400: "#45e2ab",
          500: "#16c28b",
          600: "#0a9b73",
          700: "#0a7b5f",
          800: "#0b624e",
          900: "#0b5041"
        }
      }
    }
  },
  plugins: [require("@tailwindcss/forms"), require("@tailwindcss/typography")]
} satisfies Config;
