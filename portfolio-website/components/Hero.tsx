'use client'

import { motion } from 'framer-motion'
import { ArrowDown, ChevronRight } from 'lucide-react'

const titleWords = ['Finance', 'Automation', 'Engineer']

const container = {
  hidden: {},
  show: { transition: { staggerChildren: 0.12, delayChildren: 0.3 } },
}
const wordVariant = {
  hidden: { opacity: 0, y: 40 },
  show:   { opacity: 1, y: 0, transition: { duration: 0.7, ease: [0.25, 0.46, 0.45, 0.94] } },
}

export default function Hero() {
  return (
    <section
      id="hero"
      className="relative min-h-screen flex flex-col items-center justify-center text-center overflow-hidden"
    >
      {/* Animated grid background */}
      <div className="absolute inset-0 grid-bg opacity-40 pointer-events-none" />

      {/* Gradient orbs */}
      <div
        className="absolute top-1/4 -left-40 w-[500px] h-[500px] rounded-full pointer-events-none"
        style={{
          background: 'radial-gradient(circle, rgba(0,212,255,0.08) 0%, transparent 70%)',
          animation: 'float 8s ease-in-out infinite',
        }}
      />
      <div
        className="absolute bottom-1/4 -right-40 w-[400px] h-[400px] rounded-full pointer-events-none"
        style={{
          background: 'radial-gradient(circle, rgba(124,58,237,0.10) 0%, transparent 70%)',
          animation: 'float 10s ease-in-out infinite reverse',
        }}
      />

      {/* Scanline shimmer */}
      <div className="absolute inset-0 bg-gradient-to-b from-transparent via-cyan-400/[0.012] to-transparent pointer-events-none" />

      {/* Content */}
      <div className="relative z-10 max-w-5xl mx-auto px-6">

        {/* Terminal prompt pill */}
        <motion.div
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.5, delay: 0.1 }}
          className="inline-flex items-center gap-2 px-4 py-2 rounded-full glass border border-cyan-400/20 text-xs font-mono text-cyan-400 mb-8"
        >
          <span className="w-2 h-2 rounded-full bg-cyan-400 animate-pulse" />
          python main.py --automate finance
        </motion.div>

        {/* Main title */}
        <motion.h1
          variants={container}
          initial="hidden"
          animate="show"
          className="text-5xl md:text-7xl lg:text-8xl font-bold tracking-tight mb-6"
          aria-label="Finance Automation Engineer"
        >
          {titleWords.map((word, i) => (
            <motion.span
              key={word}
              variants={wordVariant}
              className={`inline-block mr-4 last:mr-0 ${
                word === 'Automation'
                  ? 'gradient-text text-glow-cyan'
                  : 'text-white'
              }`}
            >
              {word}
            </motion.span>
          ))}
        </motion.h1>

        {/* Subtitle */}
        <motion.p
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.7, delay: 0.9 }}
          className="text-lg md:text-xl text-gray-400 max-w-2xl mx-auto mb-12 leading-relaxed"
        >
          Migrating finance teams from{' '}
          <span className="text-amber-400 font-mono">VBA</span>
          {' '}&amp;{' '}
          <span className="text-orange-400 font-mono">Alteryx</span>
          {' '}to production{' '}
          <span className="text-cyan-400 font-mono">Python</span>
          {' '}—
          one manual process at a time.
        </motion.p>

        {/* CTA buttons */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 1.1 }}
          className="flex flex-col sm:flex-row gap-4 justify-center"
        >
          <a href="#projects" className="btn-primary text-base px-8 py-4 rounded-xl shadow-cyan-sm">
            View Projects
            <ChevronRight size={18} />
          </a>
          <a href="#contact" className="btn-outline text-base px-8 py-4 rounded-xl">
            Contact Me
          </a>
        </motion.div>

        {/* Tech stack mini row */}
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ duration: 0.8, delay: 1.4 }}
          className="mt-16 flex flex-wrap gap-3 justify-center"
        >
          {['Python', 'pandas', 'Streamlit', 'SAP', 'ASC 842', 'RSU/ESPP'].map((tag) => (
            <span
              key={tag}
              className="text-xs font-mono px-3 py-1 rounded-full text-gray-500 border border-white/[0.06] bg-white/[0.02]"
            >
              {tag}
            </span>
          ))}
        </motion.div>
      </div>

      {/* Scroll indicator */}
      <motion.a
        href="#about"
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ delay: 1.8, duration: 0.6 }}
        className="absolute bottom-10 left-1/2 -translate-x-1/2 flex flex-col items-center gap-1 text-gray-600 hover:text-cyan-400 transition-colors"
      >
        <span className="text-xs font-mono tracking-widest uppercase">Scroll</span>
        <motion.div
          animate={{ y: [0, 6, 0] }}
          transition={{ duration: 1.5, repeat: Infinity, ease: 'easeInOut' }}
        >
          <ArrowDown size={16} />
        </motion.div>
      </motion.a>
    </section>
  )
}
