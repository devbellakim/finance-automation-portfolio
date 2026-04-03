'use client'

import { useRef } from 'react'
import { motion, useInView } from 'framer-motion'
import { Mail, Linkedin, Github, ArrowRight } from 'lucide-react'

const links = [
  {
    label:    'Email',
    handle:   'bellakim10@gmail.com',
    href:     'mailto:bellakim10@gmail.com',
    icon:     Mail,
    accent:   'cyan',
    desc:     'Best for project inquiries',
  },
  {
    label:    'LinkedIn',
    handle:   '/in/yourprofile',
    href:     'https://www.linkedin.com/in/bella-cpa/',
    icon:     Linkedin,
    accent:   'purple',
    desc:     'Connect professionally',
  },
  {
    label:    'GitHub',
    handle:   'github.com/yourusername',
    href:     'https://github.com/devbellakim',
    icon:     Github,
    accent:   'amber',
    desc:     'Browse the source code',
  },
] as const

const accentMap = {
  cyan:   {
    icon:   'text-cyan-400 bg-cyan-400/10 border-cyan-400/20',
    hover:  'hover:border-cyan-400/40 hover:shadow-[0_0_24px_rgba(0,212,255,0.08)]',
    arrow:  'text-cyan-400',
    handle: 'text-cyan-400',
  },
  purple: {
    icon:   'text-purple-400 bg-purple-500/10 border-purple-500/20',
    hover:  'hover:border-purple-400/40 hover:shadow-[0_0_24px_rgba(124,58,237,0.10)]',
    arrow:  'text-purple-400',
    handle: 'text-purple-400',
  },
  amber:  {
    icon:   'text-amber-400 bg-amber-400/10 border-amber-400/20',
    hover:  'hover:border-amber-400/40 hover:shadow-[0_0_24px_rgba(251,191,36,0.08)]',
    arrow:  'text-amber-400',
    handle: 'text-amber-400',
  },
}

export default function Contact() {
  const ref    = useRef(null)
  const inView = useInView(ref, { once: true, margin: '-80px' })

  return (
    <section
      id="contact"
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
            'radial-gradient(ellipse 60% 40% at 50% 60%, rgba(124,58,237,0.05) 0%, transparent 70%)',
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
          <span className="text-xs font-mono text-cyan-400 tracking-widest uppercase">05 / Contact</span>
          <div className="flex-1 h-px bg-gradient-to-r from-cyan-400/30 to-transparent" />
        </motion.div>

        <motion.h2
          initial={{ opacity: 0, y: 24 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6, delay: 0.1 }}
          className="text-4xl md:text-5xl font-bold text-white mb-4"
        >
          Get in Touch
        </motion.h2>

        <motion.p
          initial={{ opacity: 0, y: 20 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6, delay: 0.2 }}
          className="text-gray-400 text-lg max-w-xl mb-16 leading-relaxed"
        >
          Open to finance automation roles, freelance projects, and conversations
          about migrating legacy workflows to Python.
        </motion.p>

        {/* Contact cards */}
        <div className="grid md:grid-cols-3 gap-6 max-w-4xl">
          {links.map((link, i) => {
            const Icon    = link.icon
            const colours = accentMap[link.accent]

            return (
              <motion.a
                key={link.label}
                href={link.href}
                target={link.label !== 'Email' ? '_blank' : undefined}
                rel={link.label !== 'Email' ? 'noopener noreferrer' : undefined}
                initial={{ opacity: 0, y: 40 }}
                animate={inView ? { opacity: 1, y: 0 } : {}}
                transition={{ duration: 0.65, delay: 0.25 + i * 0.1 }}
                className={`glass rounded-2xl p-7 flex flex-col gap-5 group
                  border border-white/[0.06] transition-all duration-300 ${colours.hover}`}
              >
                {/* Icon */}
                <div className={`w-12 h-12 rounded-xl flex items-center justify-center border ${colours.icon}`}>
                  <Icon size={20} />
                </div>

                {/* Text */}
                <div className="flex-1">
                  <div className="text-base font-semibold text-white mb-1">{link.label}</div>
                  <div className={`text-xs font-mono mb-2 ${colours.handle}`}>{link.handle}</div>
                  <div className="text-xs text-gray-500">{link.desc}</div>
                </div>

                {/* Arrow */}
                <div className={`flex items-center gap-1 text-xs font-mono ${colours.arrow}
                  opacity-0 group-hover:opacity-100 -translate-x-1 group-hover:translate-x-0
                  transition-all duration-300`}
                >
                  Open <ArrowRight size={12} />
                </div>
              </motion.a>
            )
          })}
        </div>

        {/* Availability tag */}
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={inView ? { opacity: 1, y: 0 } : {}}
          transition={{ duration: 0.6, delay: 0.6 }}
          className="mt-14 inline-flex items-center gap-2 px-4 py-2 rounded-full glass
            border border-emerald-400/20 text-xs font-mono text-emerald-400"
        >
          <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
          Open to opportunities — based near [Ann Arbor, MI]
        </motion.div>
      </div>
    </section>
  )
}
