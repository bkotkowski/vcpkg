if {![package vsatisfies [package provide Tcl] 8.5]} {return}
package ifneeded crimp::sgi 0.2 [list ::apply {dir {
    source [file join $dir critcl-rt.tcl]
    set path [file join $dir [::critcl::runtime::MapPlatform]]
    set ext [info sharedlibextension]
    set lib [file join $path "crimp_sgi$ext"]
    load $lib Crimp_sgi
    ::critcl::runtime::Fetch $dir policy_sgi_1.tcl
    package provide crimp::sgi 0.2
}} $dir]
