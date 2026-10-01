import { useState } from 'react'
import { Link, NavLink, Route, Routes, useNavigate, useSearchParams } from 'react-router-dom'
import { ArrowDownRight, ArrowRight, ArrowUpRight, BriefcaseBusiness, MapPin, Search, Sparkles } from 'lucide-react'
import './App.css'

const opportunities = [
  { id: 1, role: 'Frontend Developer Intern', company: 'Northstar Labs', category: 'Technology', location: 'Remote', mode: 'Remote', stipend: '₹18,000 / month', duration: '3 months', mark: 'N', tone: 'mint', tags: ['React', 'JavaScript', 'CSS'] },
  { id: 2, role: 'Product Design Intern', company: 'Common Ground', category: 'Design', location: 'Bengaluru', mode: 'Hybrid', stipend: '₹15,000 / month', duration: '6 months', mark: 'C', tone: 'peach', tags: ['Figma', 'Research', 'UI'] },
  { id: 3, role: 'Digital Marketing Intern', company: 'Goodfolk Studio', category: 'Marketing', location: 'Mumbai', mode: 'On-site', stipend: '₹12,000 / month', duration: '4 months', mark: 'G', tone: 'blue', tags: ['Content', 'Social', 'SEO'] },
  { id: 4, role: 'Data Analyst Intern', company: 'Fieldnote', category: 'Technology', location: 'Remote', mode: 'Remote', stipend: '₹20,000 / month', duration: '3 months', mark: 'F', tone: 'yellow', tags: ['SQL', 'Python', 'Excel'] },
  { id: 5, role: 'People & Culture Intern', company: 'Daybreak Works', category: 'People', location: 'Delhi', mode: 'Hybrid', stipend: '₹10,000 / month', duration: '6 months', mark: 'D', tone: 'lavender', tags: ['Hiring', 'Culture', 'HR'] },
  { id: 6, role: 'Visual Content Intern', company: 'Sunday Supply', category: 'Design', location: 'Pune', mode: 'On-site', stipend: '₹14,000 / month', duration: '3 months', mark: 'S', tone: 'coral', tags: ['Photography', 'Adobe', 'Video'] },
]

function Header() {
  return <header className="site-header"><div className="header-inner">
    <Link className="brand" to="/" aria-label="DG Interns Hub home"><span className="brand-mark">dg<span>.</span></span><span className="brand-name">interns<br />hub</span></Link>
    <nav className="main-nav" aria-label="Main navigation"><NavLink to="/" end>Home</NavLink><NavLink to="/jobs">Find an internship</NavLink><NavLink to="/contact">Contact</NavLink></nav>
    <Link className="header-cta" to="/jobs">Explore roles <ArrowUpRight size={16} /></Link>
  </div></header>
}

function Footer() {
  return <footer className="site-footer"><Link className="brand footer-brand" to="/"><span className="brand-mark">dg<span>.</span></span><span className="brand-name">interns<br />hub</span></Link><p>Your first good work story starts here.</p><span className="footer-note">Made for the next generation of work.</span></footer>
}

function JobCard({ job, compact = false }) {
  return <article className={`job-card${compact ? ' job-card-compact' : ''}`}>
    <div className="job-card-top"><div className={`company-mark ${job.tone}`}>{job.mark}</div><span className="job-mode">{job.mode}</span></div>
    <p className="company-name">{job.company}</p><h3>{job.role}</h3>
    <div className="job-meta"><span><MapPin size={14} />{job.location}</span><span><BriefcaseBusiness size={14} />{job.duration}</span></div>
    {!compact && <div className="job-tags">{job.tags.map((tag) => <span key={tag}>{tag}</span>)}</div>}
    <div className="job-card-bottom"><div><small>Monthly stipend</small><strong>{job.stipend}</strong></div><a className="apply-link" href={`mailto:careers@dginternshub.com?subject=${encodeURIComponent(`Application: ${job.role}`)}`} aria-label={`Apply for ${job.role}`}><ArrowUpRight size={18} /></a></div>
  </article>
}

function HomePage() {
  const [query, setQuery] = useState('')
  const navigate = useNavigate()
  function submitSearch(event) {
    event.preventDefault()
    navigate(`/jobs${query ? `?q=${encodeURIComponent(query)}` : ''}`)
  }
  return <main>
    <section className="hero"><div className="hero-copy">
      <div className="eyebrow hero-eyebrow"><span className="eyebrow-dot" /> GOOD WORK STARTS SOMEWHERE</div>
      <h1>Your next chapter<br />starts <em>here.</em></h1>
      <p>Find internships that turn what you’re learning into what you can do. Meet teams ready to give you room to grow.</p>
      <form className="hero-search" onSubmit={submitSearch}><Search size={20} aria-hidden="true" /><input aria-label="Search internships" placeholder="Role, skill, or company" value={query} onChange={(event) => setQuery(event.target.value)} /><button type="submit">Find a role <ArrowRight size={16} /></button></form>
      <div className="hero-proof"><span className="proof-avatars"><i>R</i><i>A</i><i>M</i></span><span>Good people. First opportunities.<br /><strong>One place to find both.</strong></span></div>
    </div><div className="hero-visual" role="img" aria-label="A small creative team sharing ideas around a table"><div className="visual-caption"><span>MAKE YOUR FIRST MOVE</span><ArrowDownRight size={18} /></div><div className="visual-stamp"><Sparkles size={18} /><span>START<br />SOMETHING</span></div></div></section>
    <section className="opportunities section-wrap"><div className="section-heading"><div><span className="eyebrow">A GOOD PLACE TO BEGIN</span><h2>Open doors, right now.</h2></div><Link className="text-link" to="/jobs">See all opportunities <ArrowRight size={17} /></Link></div><div className="job-grid featured-grid">{opportunities.slice(0, 3).map((job) => <JobCard key={job.id} job={job} compact />)}</div></section>
    <section className="home-note"><div className="note-number">01 / 03</div><p>Less scrolling, more <em>starting.</em></p><Link to="/jobs" aria-label="Explore internship opportunities"><ArrowUpRight size={22} /></Link></section>
  </main>
}

function JobsPage() {
  const [searchParams, setSearchParams] = useSearchParams()
  const [query, setQuery] = useState(searchParams.get('q') ?? '')
  const [category, setCategory] = useState('All fields')
  const [mode, setMode] = useState('Any location')
  const filteredJobs = opportunities.filter((job) => {
    const matchesQuery = `${job.role} ${job.company} ${job.category} ${job.tags.join(' ')}`.toLowerCase().includes(query.toLowerCase())
    return matchesQuery && (category === 'All fields' || job.category === category) && (mode === 'Any location' || job.mode === mode)
  })
  return <main className="jobs-page"><section className="page-intro"><span className="eyebrow">YOUR SEARCH STARTS HERE</span><h1>Find work that<br /><em>moves you.</em></h1><p>Real teams. Real projects. A first step that feels like yours.</p></section>
    <section className="listing-section"><form className="filter-bar" onSubmit={(event) => { event.preventDefault(); setSearchParams(query ? { q: query } : {}) }}>
      <label className="filter-search"><Search size={18} /><input aria-label="Search by title, skill, or company" placeholder="Title, skill, or company" value={query} onChange={(event) => setQuery(event.target.value)} /></label>
      <label className="filter-select"><span className="sr-only">Field</span><select value={category} onChange={(event) => setCategory(event.target.value)}><option>All fields</option><option>Technology</option><option>Design</option><option>Marketing</option><option>People</option></select></label>
      <label className="filter-select"><span className="sr-only">Work location</span><select value={mode} onChange={(event) => setMode(event.target.value)}><option>Any location</option><option>Remote</option><option>Hybrid</option><option>On-site</option></select></label><button className="filter-submit" type="submit">Search <ArrowRight size={16} /></button>
    </form><div className="results-line"><span><strong>{filteredJobs.length}</strong> opportunities to explore</span><span>Updated this week</span></div>
    {filteredJobs.length ? <div className="job-grid">{filteredJobs.map((job) => <JobCard key={job.id} job={job} />)}</div> : <div className="empty-state"><Search size={24} /><h2>No roles found just yet.</h2><p>Try a different keyword or broaden your filters.</p><button type="button" className="text-link" onClick={() => { setQuery(''); setCategory('All fields'); setMode('Any location'); setSearchParams({}) }}>Clear filters <ArrowRight size={16} /></button></div>}
    </section></main>
}

function ContactPage() {
  const [sent, setSent] = useState(false)
  function submitMessage(event) { event.preventDefault(); setSent(true) }
  return <main className="contact-page"><section className="contact-intro"><span className="eyebrow">WE'RE ALL EARS</span><h1>Say hello.<br /><em>Mean it.</em></h1><p>Questions about a role, an idea to share, or just figuring out your next step? We’re listening.</p><a className="contact-email" href="mailto:hello@dginternshub.com">hello@dginternshub.com <ArrowUpRight size={17} /></a></section>
    <section className="contact-form-wrap">{sent ? <div className="form-success"><span className="success-icon"><Sparkles size={22} /></span><h2>Message received.</h2><p>Thanks for reaching out. We’ll be in touch soon.</p><button type="button" className="text-link" onClick={() => setSent(false)}>Send another message <ArrowRight size={16} /></button></div> : <form className="contact-form" onSubmit={submitMessage}>
      <div className="form-heading"><span className="eyebrow">DROP US A NOTE</span><h2>What’s on your mind?</h2></div><label>Your name<input required name="name" placeholder="The name your friends use" /></label><label>Email address<input required name="email" type="email" placeholder="you@example.com" /></label><label>Your message<textarea required name="message" rows="4" placeholder="Tell us a little about it..." /></label><button className="send-button" type="submit">Send message <ArrowUpRight size={17} /></button>
    </form>}</section></main>
}

function App() {
  return <div className="site-shell"><Header /><Routes><Route path="/" element={<HomePage />} /><Route path="/jobs" element={<JobsPage />} /><Route path="/contact" element={<ContactPage />} /><Route path="*" element={<HomePage />} /></Routes><Footer /></div>
}

export default App
