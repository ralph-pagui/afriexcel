/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#ecfdf5',
          100: '#d1fae5',
          200: '#a7f3d0',
          500: '#10b981',
          600: '#059669', // Vert corporate / Afrique / Excel
          700: '#047857',
          800: '#065f46',
          900: '#064e3b',
        },
        excel: {
          header: '#107c41',
          hover: '#0e6937',
          dark: '#0a4d28',
          light: '#e8f5e9'
        }
      }
    },
  },
  plugins: [],
}
