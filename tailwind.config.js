/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: {
          red: '#cf1b22',
          'red-dark': '#a5151b',
          'red-soft': '#fce8e9',
          'red-muted': '#f5c2c4',
          gray: '#50504f',
          'gray-light': '#6e6e6d',
          'gray-soft': '#f4f4f3',
          'gray-border': '#d8d8d7',
          white: '#FFFFFF',
          ink: '#020202',
        },
      },
      fontFamily: {
        sans: ['Manrope', 'Segoe UI', 'system-ui', 'sans-serif'],
      },
      boxShadow: {
        brand: '0 8px 24px -12px rgba(207, 27, 34, 0.35)',
        panel: '0 1px 3px rgba(80, 80, 79, 0.08), 0 8px 24px rgba(80, 80, 79, 0.06)',
      },
      backgroundImage: {
        'brand-mesh':
          'radial-gradient(ellipse 80% 50% at 0% 0%, rgba(207, 27, 34, 0.07), transparent 55%), radial-gradient(ellipse 60% 40% at 100% 0%, rgba(80, 80, 79, 0.06), transparent 50%), linear-gradient(180deg, #FFFFFF 0%, #f7f7f6 100%)',
      },
    },
  },
  plugins: [],
};
