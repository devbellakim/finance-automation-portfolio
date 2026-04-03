'use client'

import { useRef } from 'react'
import { motion, useInView } from 'framer-motion'
import { TrendingUp, Code2, Quote } from 'lucide-react'

const fadeUp = (delay = 0) => ({
  hidden:  { opacity: 0, y: 32 },
  show:    { opacity: 1, y: 0, transition: { duration: 0.65, delay, ease: [0.25, 0.46, 0.45, 0.94] } },
})

const financeItems = [
  { label: 'SAP ECC / S4HANA',          note: 'GL, AP, AR modules' },
  { label: 'GLSU (Excel add-in)',        note: 'Mass journal entry upload' },
  { label: 'ASC 842 Lease Accounting',  note: 'Amortization & JE prep' },
  { label: 'RSU / ESPP Reporting',      note: 'Equity comp administration' },
  { label: 'Month-End Close',           note: 'Reconciliations & accruals' },
  { label: 'Management Reporting',      note: 'CFO deck automation' },
]

const techItems = [
  { label: 'Python 3.14',    note: 'pandas · openpyxl · pptx' },
  { label: 'Streamlit',      note: 'Internal finance web apps' },
  { label: 'Alteryx',        note: 'Legacy (now migrating away)' },
  { label: 'VBA / Macros',   note: 'Legacy (now migrating away)' },
  { label: 'Git & GitHub',   note: 'Version-controlled workflows' },
  { label: 'Next.js',        note: 'This portfolio site' },
]

function Card({
  icon: Icon,
  title,
  items,
  accent,
  delay,
}: {
  icon: React.ElementType
  title: string
  items: { label: string; note: string }[]
  accent: 'cyan' | 'purple'
  delay: number
}) {
  const ref   = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  const accentClass = accent === 'cyan'
    ? 'text-cyan-400 border-cyan-400/20 bg-cyan-400/10'
    : 'text-purple-400 border-purple-500/20 bg-purple-500/10'

  return (
    <motion.div
      ref={ref}
      variants={fadeUp(delay)}
      initial="hidden"
      animate={inView ? 'show' : 'hidden'}
      className="glass rounded-2xl p-8 flex flex-col gap-6 gradient-border"
    >
      <div className="flex items-center gap-3">
        <div className={`w-10 h-10 rounded-xl flex items-center justify-center border ${accentClass}`}>
          <Icon size={18} />
        </div>
        <h3 className="text-lg font-semibold text-white">{title}</h3>
      </div>

      <ul className="flex flex-col gap-3">
        {items.map(({ label, note }) => (
          <li key={label} className="flex items-start gap-3 group">
            <span className={`mt-1 w-1.5 h-1.5 rounded-full shrink-0 ${
              accent === 'cyan' ? 'bg-cyan-400' : 'bg-purple-400'
            }`} />
            <div>
              <span className="text-sm text-gray-200 font-medium">{label}</span>
              <span className="text-xs text-gray-500 ml-2">{note}</span>
            </div>
          </li>
        ))}
      </ul>
    </motion.div>
  )
}

export default function About() {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  return (
    <section id="about" className="section-padding max-w-7xl mx-auto">
      {/* Section label */}
      <motion.div
        ref={ref}
        variants={fadeUp(0)}
        initial="hidden"
        animate={inView ? 'show' : 'hidden'}
        className="flex items-center gap-3 mb-4"
      >
        <span className="text-xs font-mono text-cyan-400 tracking-widest uppercase">01 / About</span>
        <div className="flex-1 h-px bg-gradient-to-r from-cyan-400/30 to-transparent" />
      </motion.div>

      {/* Heading */}
      <motion.h2
        variants={fadeUp(0.1)}
        initial="hidden"
        animate={inView ? 'show' : 'hidden'}
        className="text-4xl md:text-5xl font-bold text-white mb-4"
      >
        Senior Financial Analyst
        <br />
        <span className="gradient-text">turned Automation Engineer</span>
      </motion.h2>

      <motion.p
        variants={fadeUp(0.2)}
        initial="hidden"
        animate={inView ? 'show' : 'hidden'}
        className="text-gray-400 text-lg max-w-2xl mb-14 leading-relaxed"
      >
        Started as the only person on the finance team writing Python scripts.
        Ended up leading the migration of the entire department off Alteryx and VBA —
        replacing licensed tools with maintainable, version-controlled Python pipelines.
      </motion.p>

      {/* Highlight quote */}
      <motion.div
        variants={fadeUp(0.25)}
        initial="hidden"
        animate={inView ? 'show' : 'hidden'}
        className="glass-cyan rounded-2xl p-6 mb-14 flex gap-4 items-start max-w-3xl"
      >
        <Quote size={20} className="text-cyan-400 shrink-0 mt-0.5" />
        <p className="text-cyan-100 text-sm leading-relaxed font-medium">
          &ldquo;Only Python user on the team → identified tooling gaps → built 4 production automation apps →
          led Alteryx licence migration, saving the department significant annual tooling cost.&rdquo;
        </p>
      </motion.div>

      {/* Two-column cards */}
      <div className="grid md:grid-cols-2 gap-6">
        <Card
          icon={TrendingUp}
          title="Finance Background"
          items={financeItems}
          accent="cyan"
          delay={0.3}
        />
        <Card
          icon={Code2}
          title="Tech Skills"
          items={techItems}
          accent="purple"
          delay={0.4}
        />
      </div>
    </section>
  )
}
