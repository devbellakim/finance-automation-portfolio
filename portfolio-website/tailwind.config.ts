import type { Config } from 'tailwindcss'

const config: Config = {
  content: [
    './pages/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
    './app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        background:  '#0A0A0F',
        surface:     '#0F0F1A',
        'surface-2': '#141428',
        cyan: {
          400: '#00D4FF',
          500: '#00BFDF',
          600: '#009DBF',
        },
        purple: {
          500: '#7C3AED',
          600: '#6D28D9',
          700: '#5B21B6',
        },
      },
      fontFamily: {
        sans: ['var(--font-inter)', 'system-ui', 'sans-serif'],
        mono: ['var(--font-jetbrains)', 'monospace'],
      },
      backgroundImage: {
        'grid-cyan': `
          linear-gradient(rgba(0,212,255,0.04) 1px, transparent 1px),
          linear-gradient(90deg, rgba(0,212,255,0.04) 1px, transparent 1px)
        `,
      },
      backgroundSize: {
        'grid': '60px 60px',
      },
      animation: {
        'grid-drift': 'gridDrift 25s linear infinite',
        'pulse-slow':  'pulse 4s cubic-bezier(0.4,0,0.6,1) infinite',
        'pulse-slower':'pulse 6s cubic-bezier(0.4,0,0.6,1) infinite',
        'float':       'float 6s ease-in-out infinite',
        'glow':        'glow 2s ease-in-out infinite alternate',
      },
      keyframes: {
        gridDrift: {
          '0%':   { backgroundPosition: '0px 0px' },
          '100%': { backgroundPosition: '60px 60px' },
        },
        float: {
          '0%, 100%': { transform: 'translateY(0px)' },
          '50%':      { transform: 'translateY(-20px)' },
        },
        glow: {
          from: { boxShadow: '0 0 20px rgba(0,212,255,0.1)' },
          to:   { boxShadow: '0 0 40px rgba(0,212,255,0.3)' },
        },
      },
      boxShadow: {
        'cyan-sm':  '0 0 20px rgba(0,212,255,0.15)',
        'cyan-md':  '0 0 40px rgba(0,212,255,0.25)',
        'purple-sm':'0 0 20px rgba(124,58,237,0.15)',
      },
    },
  },
  plugins: [],
}

export default config
