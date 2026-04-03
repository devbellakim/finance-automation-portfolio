import Nav     from '@/components/Nav'
import Hero    from '@/components/Hero'
import About   from '@/components/About'
import Skills  from '@/components/Skills'
import Projects from '@/components/Projects'
import Impact  from '@/components/Impact'
import Contact from '@/components/Contact'

export default function Home() {
  return (
    <main className="relative overflow-x-hidden">
      <Nav />
      <Hero />
      <About />
      <Skills />
      <Projects />
      <Impact />
      <Contact />
      <footer className="text-center py-8 text-gray-600 text-sm font-mono border-t border-white/5">
        <span className="text-cyan-400/60">{'>'}</span>
        {' '}Built with Next.js 14 · Tailwind · Framer Motion
        {' '}·{' '}
        <span className="text-cyan-400/60">{'</'}</span>
        Finance Automation
        <span className="text-cyan-400/60">{'>'}</span>
      </footer>
    </main>
  )
}
