'use client'

import { useRef } from 'react'
import { motion, useInView } from 'framer-motion'

const groups = [
  {
    label:  'Finance Tools',
    accent: 'amber' as const,
    description: 'Enterprise systems and reporting platforms',
    items: [
      { name: 'SAP ECC / S4HANA', icon: '⬡' },
      { name: 'Alteryx',          icon: '⬡' },
      { name: 'Microsoft Excel',  icon: '⬡' },
      { name: 'VBA / Macros',     icon: '⬡' },
      { name: 'GLSU',             icon: '⬡' },
      { name: 'PowerPoint',       icon: '⬡' },
      { name: 'Power BI',         icon: '⬡' },
    ],
  },
  {
    label:  'Python Stack',
    accent: 'cyan' as const,
    description: 'Libraries powering the automations',
    items: [
      { name: 'Python 3.14',  icon: '◈' },
      { name: 'pandas',       icon: '◈' },
      { name: 'openpyxl',     icon: '◈' },
      { name: 'python-pptx',  icon: '◈' },
      { name: 'Streamlit',    icon: '◈' },
      { name: 'numpy',        icon: '◈' },
      { name: 'Plotly',       icon: '◈' },
    ],
  },
  {
    label:  'Dev Tools',
    accent: 'purple' as const,
    description: 'Infrastructure and developer workflow',
    items: [
      { name: 'Git & GitHub',  icon: '◇' },
      { name: 'VS Code',       icon: '◇' },
      { name: 'Next.js 14',    icon: '◇' },
      { name: 'TypeScript',    icon: '◇' },
      { name: 'Tailwind CSS',  icon: '◇' },
      { name: 'Framer Motion', icon: '◇' },
    ],
  },
]

const badgeColor = {
  amber:  'bg-amber-400/10 text-amber-400 border-amber-400/25 hover:bg-amber-400/20',
  cyan:   'bg-cyan-400/10 text-cyan-400 border-cyan-400/25 hover:bg-cyan-400/20',
  purple: 'bg-purple-500/10 text-purple-400 border-purple-500/25 hover:bg-purple-400/20',
}

const headerColor = {
  amber:  'text-amber-400',
  cyan:   'text-cyan-400',
  purple: 'text-purple-400',
}

const barColor = {
  amber:  'from-amber-400/60 to-amber-400/10',
  cyan:   'from-cyan-400/60 to-cyan-400/10',
  purple: 'from-purple-400/60 to-purple-400/10',
}

export default function Skills() {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  return (
    <section id="skills" className="section-padding max-w-7xl mx-auto">
      {/* Header */}
      <motion.div
        ref={ref}
        initial={{ opacity: 0, y: 24 }}
        animate={inView ? { opacity: 1, y: 0 } : {}}
        transition={{ duration: 0.6 }}
        className="flex items-center gap-3 mb-4"
      >
        <span className="text-xs font-mono text-cyan-400 tracking-widest uppercase">02 / Skills</span>
        <div className="flex-1 h-px bg-gradient-to-r from-cyan-400/30 to-transparent" />
      </motion.div>

      <motion.h2
        initial={{ opacity: 0, y: 24 }}
        animate={inView ? { opacity: 1, y: 0 } : {}}
        transition={{ duration: 0.6, delay: 0.1 }}
        className="text-4xl md:text-5xl font-bold text-white mb-14"
      >
        Tools &amp; Technologies
      </motion.h2>

      {/* Skill groups */}
      <div className="grid md:grid-cols-3 gap-8">
        {groups.map((group, gi) => (
          <motion.div
            key={group.label}
            initial={{ opacity: 0, y: 40 }}
            animate={inView ? { opacity: 1, y: 0 } : {}}
            transition={{ duration: 0.65, delay: 0.15 + gi * 0.12 }}
            className="glass rounded-2xl p-7 flex flex-col gap-5"
          >
            {/* Group header */}
            <div>
              <div
                className={`h-0.5 w-10 rounded mb-4 bg-gradient-to-r ${barColor[group.accent]}`}
              />
              <h3 className={`text-base font-semibold ${headerColor[group.accent]}`}>
                {group.label}
              </h3>
              <p className="text-xs text-gray-500 mt-1">{group.description}</p>
            </div>

            {/* Badges */}
            <div className="flex flex-wrap gap-2">
              {group.items.map((item, ii) => (
                <motion.span
                  key={item.name}
                  initial={{ opacity: 0, scale: 0.85 }}
                  animate={inView ? { opacity: 1, scale: 1 } : {}}
                  transition={{ duration: 0.3, delay: 0.3 + gi * 0.1 + ii * 0.04 }}
                  className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-mono
                    font-medium border cursor-default transition-colors duration-200
                    ${badgeColor[group.accent]}`}
                >
                  <span className="opacity-60 text-[10px]">{item.icon}</span>
                  {item.name}
                </motion.span>
              ))}
            </div>
          </motion.div>
        ))}
      </div>
    </section>
  )
}
