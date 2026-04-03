'use client'

import { useRef } from 'react'
import { motion, useInView } from 'framer-motion'
import { Play, ExternalLink, Github, Star } from 'lucide-react'

interface Project {
  id:       string
  title:    string
  oneliner: string
  problem:  string
  solution: string
  impact:   string
  stack:    string[]
  featured?: boolean
}

const projects: Project[] = [
  {
    id:       'p2',
    title:    'ASC 842 Lease Automation',
    oneliner: 'Quarterly lease variance & journal entry — 3-platform migration',
    problem:  'Manual quarterly lease variance analysis and journal entry preparation across two legacy platforms (Alteryx + VBA), prone to errors during month-end close.',
    solution: 'Full Alteryx → VBA → Python Streamlit migration. Compares Q-over-Q Lease Harbor reports, auto-generates ASC 842-compliant journal entries by lease type.',
    impact:   'Eliminated 3-platform dependency, reduced JE prep time from hours to minutes, and cut Alteryx licence costs.',
    stack:    ['Python', 'pandas', 'Streamlit', 'openpyxl'],
    featured: true,
  },
  {
    id:       'p1',
    title:    'SAP Report Automation',
    oneliner: 'GL export → formatted management report pipeline',
    problem:  'Finance team spent hours manually pivot-tabling and VLOOKUPing raw SAP GL exports each month-end.',
    solution: 'Python pipeline: loads SAP export, maps GL codes, pivots by account/cost centre, generates formatted Excel report and email-ready summary.',
    impact:   'Eliminated manual formatting; output is reproducible, version-controlled, and email-ready in seconds.',
    stack:    ['Python', 'pandas', 'Streamlit', 'openpyxl'],
  },
  {
    id:       'p3',
    title:    'Excel → PowerPoint Automation',
    oneliner: 'Auto-refresh CFO deck from live Excel ranges',
    problem:  'Analysts spent 1–2 hours per cycle copying financial charts and tables from Excel into a PowerPoint deck, manually reformatting everything.',
    solution: 'python-pptx pipeline reads named Excel ranges and injects them into a slide template at fixed positions — one command refreshes the entire deck.',
    impact:   'Deck refresh time: 2 hours → under 1 minute. Eliminated copy-paste errors across revision cycles.',
    stack:    ['Python', 'python-pptx', 'openpyxl', 'Streamlit'],
  },
  {
    id:       'p4',
    title:    'RSU / ESPP Equity Tracker',
    oneliner: 'Messy Fidelity export → clean equity report pipeline',
    problem:  'Fidelity equity exports have sparse rows (Employee_ID blank for rows 2–N per employee). Multi-source join against HR data was done manually in Excel.',
    solution: 'Replicates the Alteryx workflow: forward-fills sparse columns, joins employee reference table (VLOOKUP equivalent), splits RSU vs ESPP, generates summary reports.',
    impact:   'Fully automated equity reporting pipeline; replaced a 3-file manual process with a single Streamlit app.',
    stack:    ['Python', 'pandas', 'Streamlit', 'openpyxl'],
  },
]

function VideoPlaceholder({ featured = false }: { featured?: boolean }) {
  return (
    <div className={`relative overflow-hidden rounded-xl flex items-center justify-center bg-gradient-to-br from-surface to-surface-2 border border-white/[0.06] group-hover:border-cyan-400/20 transition-colors ${
      featured ? 'aspect-video' : 'aspect-video'
    }`}>
      {/* Fake terminal lines */}
      <div className="absolute inset-0 p-4 flex flex-col gap-1.5 font-mono text-[10px] text-gray-700 select-none">
        {['$ python src/app.py', '> Loading data...', '> Processing 500 rows', '> Building report...', '✓ Done — report saved'].map((line) => (
          <div key={line} className="truncate">{line}</div>
        ))}
      </div>
      {/* Play button overlay */}
      <div className="relative z-10 flex flex-col items-center gap-2">
        <div className="w-14 h-14 rounded-full glass border border-cyan-400/30 flex items-center justify-center group-hover:bg-cyan-400/10 group-hover:border-cyan-400/60 transition-all duration-300">
          <Play size={20} className="text-cyan-400 ml-0.5" fill="rgba(0,212,255,0.6)" />
        </div>
        <span className="text-xs text-gray-500 font-mono">Watch Demo</span>
      </div>
    </div>
  )
}

function StackBadge({ label }: { label: string }) {
  return (
    <span className="badge-cyan text-[11px]">
      {label}
    </span>
  )
}

function ActionButtons() {
  return (
    <div className="flex gap-3 flex-wrap mt-6">
      <a href="#" className="btn-primary text-xs px-4 py-2">
        <Play size={13} />
        Watch Demo
      </a>
      <a href="#" className="btn-outline text-xs px-4 py-2">
        <ExternalLink size={13} />
        Live App
      </a>
      <a href="#" className="btn-outline text-xs px-4 py-2">
        <Github size={13} />
        GitHub
      </a>
    </div>
  )
}

function FeaturedCard({ project }: { project: Project }) {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-60px' })

  return (
    <motion.div
      ref={ref}
      initial={{ opacity: 0, y: 40 }}
      animate={inView ? { opacity: 1, y: 0 } : {}}
      transition={{ duration: 0.7, ease: [0.25, 0.46, 0.45, 0.94] }}
      className="glass rounded-2xl p-8 mb-8 grid md:grid-cols-2 gap-8 group
                 hover:border-cyan-400/20 hover:shadow-cyan-sm transition-all duration-300 gradient-border"
    >
      {/* Left: video */}
      <div className="order-2 md:order-1">
        <VideoPlaceholder featured />
      </div>

      {/* Right: content */}
      <div className="order-1 md:order-2 flex flex-col justify-between">
        <div>
          {/* Featured badge */}
          <div className="flex items-center gap-2 mb-4">
            <Star size={13} className="text-amber-400" fill="rgba(251,191,36,0.8)" />
            <span className="text-xs font-mono text-amber-400 uppercase tracking-widest">Featured Project</span>
          </div>

          <h3 className="text-2xl font-bold text-white mb-2 group-hover:text-cyan-100 transition-colors">
            {project.title}
          </h3>
          <p className="text-sm text-gray-400 mb-6">{project.oneliner}</p>

          <div className="flex flex-col gap-3">
            {(
              [
                { label: 'Problem',  text: project.problem,  color: 'text-red-400' },
                { label: 'Solution', text: project.solution, color: 'text-cyan-400' },
                { label: 'Impact',   text: project.impact,   color: 'text-emerald-400' },
              ] as { label: string; text: string; color: string }[]
            ).map(({ label, text, color }) => (
              <div key={label} className="flex gap-3 text-sm">
                <span className={`font-mono font-semibold ${color} shrink-0 w-16 text-xs mt-0.5`}>
                  {label}
                </span>
                <span className="text-gray-300 leading-relaxed">{text}</span>
              </div>
            ))}
          </div>
        </div>

        <div>
          <div className="flex flex-wrap gap-2 mt-6 mb-1">
            {project.stack.map((s) => <StackBadge key={s} label={s} />)}
          </div>
          <ActionButtons />
        </div>
      </div>
    </motion.div>
  )
}

function ProjectCard({ project, delay }: { project: Project; delay: number }) {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-60px' })

  return (
    <motion.div
      ref={ref}
      initial={{ opacity: 0, y: 40 }}
      animate={inView ? { opacity: 1, y: 0 } : {}}
      transition={{ duration: 0.65, delay, ease: [0.25, 0.46, 0.45, 0.94] }}
      className="glass rounded-2xl p-6 flex flex-col group
                 hover:border-cyan-400/15 hover:shadow-cyan-sm
                 hover:-translate-y-1 transition-all duration-300"
    >
      <VideoPlaceholder />

      <div className="mt-5 flex flex-col flex-1">
        <h3 className="text-lg font-bold text-white mb-1 group-hover:text-cyan-100 transition-colors">
          {project.title}
        </h3>
        <p className="text-xs text-gray-500 mb-4">{project.oneliner}</p>

        <div className="flex flex-col gap-2.5 flex-1">
          {(
            [
              { label: 'Problem',  text: project.problem,  color: 'text-red-400' },
              { label: 'Solution', text: project.solution, color: 'text-cyan-400' },
              { label: 'Impact',   text: project.impact,   color: 'text-emerald-400' },
            ] as { label: string; text: string; color: string }[]
          ).map(({ label, text, color }) => (
            <div key={label} className="flex gap-2 text-xs leading-relaxed">
              <span className={`font-mono font-semibold ${color} shrink-0 w-14`}>
                {label}
              </span>
              <span className="text-gray-400">{text}</span>
            </div>
          ))}
        </div>

        <div className="flex flex-wrap gap-1.5 mt-4 mb-4">
          {project.stack.map((s) => <StackBadge key={s} label={s} />)}
        </div>

        <div className="flex gap-2 flex-wrap">
          <a href="#" className="btn-primary text-xs px-3 py-1.5 gap-1.5">
            <Play size={11} /> Demo
          </a>
          <a href="#" className="btn-outline text-xs px-3 py-1.5 gap-1.5">
            <ExternalLink size={11} /> App
          </a>
          <a href="#" className="btn-outline text-xs px-3 py-1.5 gap-1.5">
            <Github size={11} />
          </a>
        </div>
      </div>
    </motion.div>
  )
}

export default function Projects() {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  const featured = projects.find((p) => p.featured)!
  const rest     = projects.filter((p) => !p.featured)

  return (
    <section id="projects" className="section-padding max-w-7xl mx-auto">
      {/* Header */}
      <motion.div
        ref={ref}
        initial={{ opacity: 0, y: 24 }}
        animate={inView ? { opacity: 1, y: 0 } : {}}
        transition={{ duration: 0.6 }}
        className="flex items-center gap-3 mb-4"
      >
        <span className="text-xs font-mono text-cyan-400 tracking-widest uppercase">03 / Projects</span>
        <div className="flex-1 h-px bg-gradient-to-r from-cyan-400/30 to-transparent" />
      </motion.div>

      <motion.h2
        initial={{ opacity: 0, y: 24 }}
        animate={inView ? { opacity: 1, y: 0 } : {}}
        transition={{ duration: 0.6, delay: 0.1 }}
        className="text-4xl md:text-5xl font-bold text-white mb-14"
      >
        Finance Automation{' '}
        <span className="gradient-text">Projects</span>
      </motion.h2>

      {/* Featured */}
      <FeaturedCard project={featured} />

      {/* Grid */}
      <div className="grid md:grid-cols-3 gap-6">
        {rest.map((project, i) => (
          <ProjectCard key={project.id} project={project} delay={i * 0.1} />
        ))}
      </div>
    </section>
  )
}
