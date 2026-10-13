import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarGroup,
  SidebarGroupLabel,
  SidebarHeader,
  SidebarGroupContent,
  SidebarMenu,
  SidebarMenuItem,
  SidebarMenuButton,
  useSidebar,
} from "@/components/ui/sidebar"
import { Button } from "@/components/ui/button"
import { CameraIcon, FileIcon, KeyIcon, LogoutIcon } from "@hugeicons/core-free-icons"
import { HugeiconsIcon } from "@hugeicons/react"
import { Link } from "react-router-dom"
import { useNavigate } from "react-router-dom"
import { useAuthStore } from "./store/auth-store"
import { useQueryClient } from "@tanstack/react-query"
import { logout } from "./lib/auth-api"
import { Popover, PopoverTrigger, PopoverContent, PopoverHeader, PopoverTitle, PopoverDescription } from "@/components/ui/popover"
export function AppSidebar() {
  const navigate = useNavigate()
  const queryClient = useQueryClient()
  const { utente } = useAuthStore()
  const firstTowLetters = utente?.email?.slice(0, 2) ?? utente?.username?.slice(0, 2)

  const { setOpenMobile } = useSidebar()
  const { jwtPayload } = useAuthStore()
  const isAdmin = jwtPayload?.ruolo === "admin"
  const menuItems = [
    { title: "Registra Copie ", url: "/registra-copie", icon: CameraIcon, isProtected: false },
    { title: "Gestione Docenti", url: "/gestione-docenti", icon: FileIcon, isProtected: true },
    { title: "Gestione Utenze", url: "/gestione-utenze", icon: KeyIcon, isProtected: true },
    { title: "Visualizza Registrazioni", url: "/visualizza-registrazioni", icon: FileIcon, isProtected: true },
  ]
  const handleLogout = async () => {
    // La funzione logout gestisce già tutto: chiamata API, pulizia store e cache
    // Non serve try-catch perché logout() gestisce gli errori internamente
    // e pulisce sempre lo store locale (necessario per permettere l'accesso a /login)
    await logout(queryClient)
    navigate('/login')
  }
  return (
    <Sidebar  >
      <SidebarHeader className="bg-sidebar-primary">

        <SidebarGroup>
          <SidebarGroupLabel className="">
            <h1 className="text-xl  font-bold whitespace-nowrap text-sidebar-accent ">ScriptaScuola</h1>
          </SidebarGroupLabel>
        </SidebarGroup>
      </SidebarHeader>
      <SidebarContent className="bg-sidebar-primary">
        <SidebarGroup>
          <SidebarGroupContent>
            <SidebarMenu>
              {menuItems.filter((item) => !item.isProtected || isAdmin).map((item) => (
                <Link key={item.title} to={item.url} className="" onClick={() => setOpenMobile(false)}>
                  <SidebarMenuItem>
                    <SidebarMenuButton className="rounded-none h-12 text-sidebar-border hover:text-primary w-full justify-between gap-2 " >

                      <p className="font-semibold">{item.title}</p>
                      <HugeiconsIcon icon={item.icon} strokeWidth={2} />

                    </SidebarMenuButton>
                  </SidebarMenuItem>
                </Link>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>
      <SidebarFooter className="bg-sidebar-primary">
      
       
            <Popover >
              <PopoverTrigger asChild className="hover:bg-secondary/20 transition-all duration-200 ease-linear p-2 w-full">
                <div className="w-full flex items-center gap-2 cursor-pointer ">
                  <p className="text-secondary font-bold text-sm p-2 bg-secondary/20 rounded-full ">{firstTowLetters}</p>
                  <p className="text-secondary text-sm">{utente?.email ?? utente?.username}</p>
                </div>
   
              </PopoverTrigger>
              <PopoverContent className="" align="start"  sideOffset={10}>
                <PopoverHeader>
                  <PopoverTitle className="flex items-center gap-2 ">
                    <p className="text-primary font-bold text-xs p-3 bg-sidebar-accent rounded-full">{firstTowLetters}</p>
                    <p className="text-primary text-xs">{utente?.email ?? utente?.username}</p>
                  </PopoverTitle>
                  <PopoverDescription className="flex flex-col gap-2">
                  <Button variant="secondary" asChild>
                    <Link to="/profilo" onClick={() => setOpenMobile(false)}>Visualizza Profilo</Link>
                  </Button>

                    <Button variant="ghost" className="" onClick={handleLogout}> <span className="text-destructive flex items-center gap-2">Esci <HugeiconsIcon icon={LogoutIcon} className="text-destructive size-4" /></span></Button>
                  </PopoverDescription>
                </PopoverHeader>
              </PopoverContent>
            </Popover>
          

     
      </SidebarFooter>
    </Sidebar>
  )
}