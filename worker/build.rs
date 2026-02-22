fn main() {
    // Embed Windows application manifest and icon
    if std::env::var( "CARGO_CFG_TARGET_OS" ).unwrap_or_default() == "windows" {
        let mut res = winres::WindowsResource::new();
        res.set_icon( "assets/icon.ico" );
        // Use GUI subsystem so no console window appears
        res.set_manifest(
            r#"
            <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
                <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
                    <security>
                        <requestedPrivileges>
                            <requestedExecutionLevel level="asInvoker" uiAccess="false"/>
                        </requestedPrivileges>
                    </security>
                </trustInfo>
            </assembly>
            "#,
        );
        // Only embed icon if it exists (don't fail during development)
        if std::path::Path::new( "assets/icon.ico" ).exists() {
            res.compile().expect( "Failed to compile Windows resources" );
        } else {
            println!( "cargo:warning=assets/icon.ico not found, skipping resource embedding" );
        }
    }
}
