import { useState, useEffect } from "react";
import { Link } from "react-router-dom";
import { motion } from "framer-motion";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { HugeiconsIcon } from "@hugeicons/react";
import { MenuIcon } from "@hugeicons/core-free-icons";
import { Separator } from "../ui/separator";

// Sezioni della landing: costanti, quindi fuori dal componente.
const NAV_SECTIONS = [
  { id: "benefits", label: "Perché ScriptaScuola" },
  { id: "security", label: "Dati e sicurezza" },
  { id: "how-it-works", label: "Come funziona" },
  { id: "cta", label: "Inizia ora" },
];

export default function LandingHeader() {
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false)
  const [activeSection, setActiveSection] = useState<string | null>(null);
  useEffect(() => {
    const sectionsElements = NAV_SECTIONS.map((section)=> {
      const el = document.getElementById(section.id);
      return el ? {id: section.id, el} : null;
    }).filter(Boolean) as {id: string, el: HTMLElement}[];
    if(!sectionsElements.length)  return

    const observer = new IntersectionObserver(
      (entries) => {
        const visibleEntries = entries 
        .filter((entry) => entry.isIntersecting )
        if(visibleEntries[0]) {
          const id = visibleEntries[0].target.id;
          setActiveSection(id)
        }
      },
      {
        root: null,
        threshold: [0.3, 0.5, 0.7]
      }
    )
    sectionsElements.forEach(({el}) => {
      observer.observe(el);
    })
    return () => {
    
      observer.disconnect();
    } 
  }, []);

  const scrollToSection = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ behavior: "smooth", block: "start" });
    }
    setIsMobileMenuOpen(false);
  };

  return (
    <motion.header
      initial={{ opacity: 0, y: -20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5 }}
      className={cn(
        "sticky top-0 z-50 w-full border-b bg-background/95 backdrop-blur supports-backdrop-filter:bg-background/60 transition-all shadow-md",
      )}
    >
      <nav className="container mx-auto px-4 h-16 flex items-center justify-between">
        {/* Logo */}
        <button onClick={() => scrollToSection("hero")} className="flex items-center gap-2 cursor-pointer">
          <img
            src="/logo.svg"
            alt="ScriptaScuola Logo"
            className="h-8 w-8 object-contain"
          />
          <span className="text-xl font-semibold text-primary">ScriptaScuola</span>
        </button>

        {/* Navigation */}
        <nav className="hidden md:flex items-center gap-4 lg:gap-6">
          {NAV_SECTIONS.map((section) => (
            <button
              key={section.id}
              onClick={() => scrollToSection(section.id)}
              className={cn("text-sm font-medium text-muted-foreground hover:text-foreground transition-all cursor-pointer", activeSection === section.id && "text-primary font-bold")}
            >
              {section.label}
            </button>
          ))}
        </nav>

        {/* CTA Buttons */}
        <div className="hidden md:flex items-center gap-3">
          <Button
            variant="ghost"
            asChild
            className="hidden sm:flex"
          >
            <Link to="/login">Accedi</Link>
          </Button>
          <Button asChild>
            <Link to="/register">Registrati</Link>
          </Button>
        </div>
        <div className="md:hidden">
          <Button variant="ghost" asChild className="p-2" size="icon" onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}>
            <HugeiconsIcon icon={MenuIcon} className="text-primary" />
          </Button>
        </div>
      </nav>
      {isMobileMenuOpen && <div id="mobile-menu">
        <nav className="md:hidden flex flex-col items-center gap-4 py-4 shadow-md">
          {NAV_SECTIONS.map((section) => (
            <button
              key={section.id}
              onClick={() => scrollToSection(section.id)}
              className={cn("text-sm font-medium text-muted-foreground hover:text-foreground transition-all cursor-pointer", activeSection === section.id && "text-primary font-bold")}
            >
              {section.label}
            </button>
          ))}
          <Separator />
          <div className="flex flex-col gap-2 items-center">
          <Button variant="default" asChild className="w-fit">
            <Link to="/register">Registrati</Link>
          </Button>
          <Button variant="outline" asChild className="w-fit">
            <Link to="/login">Hai già un account? Accedi</Link>
          </Button>
          </div>
        </nav>
      </div>}
    </motion.header>
  )
}
