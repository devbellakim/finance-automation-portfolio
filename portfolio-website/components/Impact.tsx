'use client'

import { useRef, useEffect, useState } from 'react'
import { motion, useInView, animate } from 'framer-motion'

interface Metric {
  prefix?:  string
  value:    number
  suffix:   string
  label:    string
  sublabel: string
  accent:   'cyan' | 'purple' | 'amber' | 'emerald'
}

const metrics: Metric[] = [
  {
    value:    4,
    suffix:   '',
    label:    'Automations Built',
    sublabel: 'End-to-end Python pipelines',
    accent:   'cyan',
  },
  {
    value:    3,
    suffix:   '',
    label:    'Platforms Migrated',
    sublabel: 'Alteryx + VBA → Python',
    accent:   'purple',
  },
  {
    value:    100,
    suffix:   '%',
    label:    'Error Reduction',
    sublabel: 'Manual processes eliminated',
    accent:   'emerald',
  },
  {
    prefix:   '∞',
    value:    0,
    suffix:   '',
    label:    'Revision Cycles',
    sublabel: 'PPT deck auto-refreshes',
    accent:   'amber',
  },
]

const accentMap = {
  cyan:    { text: 'text-cyan-400',    glow: 'shadow-cyan-sm',    bar: 'from-cyan-400/60' },
  purple:  { text: 'text-purple-400',  glow: 'shadow-purple-sm',  bar: 'from-purple-400/60' },
  emerald: { text: 'text-emerald-400', glow: 'shadow-[0_0_20px_rgba(52,211,153,0.15)]', bar: 'from-emerald-400/60' },
  amber:   { text: 'text-amber-400',   glow: 'shadow-[0_0_20px_rgba(251,191,36,0.15)]', bar: 'from-amber-400/60' },
}

function Counter({ metric }: { metric: Metric }) {
  const ref      = useRef<HTMLSpanElement>(null)
  const isInView = useInView(ref, { once: true, margin: '-100px' })
  const [display, setDisplay] = useState(metric.prefix ? metric.prefix : '0')

  useEffect(() => {
    if (!isInView || metric.prefix) return
    const controls = animate(0, metric.value, {
      duration: 2,
      ease: 'easeOut',
      onUpdate: (v) => setDisplay(String(Math.round(v))),
    })
    return controls.stop
  }, [isInView, metric.value, metric.prefix])

  const { text, glow } = accentMap[metric.accent]

  return (
    <div className={`glass rounded-2xl p-8 text-center group hover:${glow} transition-all duration-300`}>
      <div className={`text-6xl md:text-7xl font-bold font-mono ${text} mb-2`}>
        <span ref={ref}>{display}</span>
        {metric.suffix && <span>{metric.suffix}</span>}
      </div>
      <div className="text-lg font-semibold text-white mb-1">{metric.label}</div>
      <div className="text-sm text-gray-500">{metric.sublabel}</div>
    </div>
  )
}

export default function Impact() {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  return (
    <section
      id="impact"
      className="relative section-padding overflow-hidden"
      style={{
        background: 'linear-gradient(180deg, #0A0A0F 0%, #0D0D1A 50%, #0A0A0F 100%)',
      }}
    >
      {/* Background glow */}
      <div
        className="absolute inset-0 pointer-events-none"
        style={{
          background:
            'radial-gradient(ellipse 80% 50% at 50% 50%, rgba(0,212,255,0.04) 0%, transparent 70%)',
        }}
      />

      <div className="relative z-10 max-w-7xl mx-auto">
        {/* Header */}
        <motion.div
          ref={ref}
          initial={{ opacity: 0, y: 24 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6 }}
          className="flex items-center gap-3 mb-4"
        >
          <span className="text-xs font-mono text-cyan-400 tracking-widest uppercase">04 / Impact</span>
          <div className="flex-1 h-px bg-gradient-to-r from-cyan-400/30 to-transparent" />
        </motion.div>

        <motion.h2
          initial={{ opacity: 0, y: 24 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6, delay: 0.1 }}
          className="text-4xl md:text-5xl font-bold text-white mb-4"
        >
          By the Numbers
        </motion.h2>

        <motion.p
          initial={{ opacity: 0, y: 20 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6, delay: 0.2 }}
          className="text-gray-400 text-lg max-w-xl mb-16"
        >
          Real outcomes from migrating manual finance workflows to automated Python pipelines.
        </motion.p>

        {/* Metric grid */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-6">
          {metrics.map((metric, i) => (
            <motion.div
              key={metric.label}
              initial={{ opacity: 0, y: 40, scale: 0.95 }}
              animate={inView ? { opacity: 1, y: 0, scale: 1 } : {}}
              transition={{ duration: 0.65, delay: 0.2 + i * 0.1 }}
            >
              <Counter metric={metric} />
            </motion.div>
          ))}
        </div>

        {/* Timeline strip */}
        <motion.div
          initial={{ opacity: 0, y: 24 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.65, delay: 0.6 }}
          className="mt-20 glass rounded-2xl p-8"
        >
          <h3 className="text-sm font-mono text-gray-500 mb-8 uppercase tracking-widest">
            Migration Timeline
          </h3>
          <div className="relative">
            {/* Line */}
            <div className="absolute top-3 left-0 right-0 h-px bg-gradient-to-r from-cyan-400/40 via-purple-400/40 to-cyan-400/10" />

            <div className="relative grid grid-cols-2 md:grid-cols-4 gap-8">
              {[
                { phase: 'Phase 1', label: 'SAP Report',  tool: 'VBA → Python',     color: 'text-cyan-400' },
                { phase: 'Phase 2', label: 'Lease Auto',  tool: 'Alteryx → Python',  color: 'text-purple-400' },
                { phase: 'Phase 3', label: 'Excel → PPT', tool: 'Manual → python-pptx', color: 'text-amber-400' },
                { phase: 'Phase 4', label: 'Equity',      tool: 'Alteryx → Python',  color: 'text-emerald-400' },
              ].map((step) => (
                <div key={step.phase} className="pt-8">
                  <div className={`w-3 h-3 rounded-full absolute top-1.5 -translate-y-1/2 ${
                    step.color.replace('text-', 'bg-')
                  } shadow-cyan-sm`} />
                  <div className="text-[10px] font-mono text-gray-600 mb-1">{step.phase}</div>
                  <div className="text-sm font-semibold text-white">{step.label}</div>
                  <div className={`text-xs font-mono mt-1 ${step.color}`}>{step.tool}</div>
                </div>
              ))}
            </div>
          </div>
        </motion.div>
      </div>
    </section>
  )
}
