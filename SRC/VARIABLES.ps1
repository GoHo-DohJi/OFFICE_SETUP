### CONST ###
$OFFICE_DIR                  = "$env:PROGRAMFILES\Microsoft Office"
$CONFIGURATION_TEMPLATE_PATH = "$PSScriptRoot\CONFIGURATION_TEMPLATE.xml"
$CONFIGURATION_PATH          = "$PSScriptRoot\CONFIGURATION.xml"
$DOWNLOAD_URL                = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img"



# Fraktur Text Art ( https://fontgenro.com/ascii-text-generator ) #
$LOGO = @"
                                   .                         
               oec :     oec :    @88>                       
       u.     @88888    @88888    %8P                        
 ...ue888b    8"*88%    8"*88%     .          .        .u    
 888R Y888r   8b.       8b.      .@88u   .udR88N    ud8888.  
 888R I888>  u888888>  u888888> '"888E` <888'888k :888'8888. 
 888R I888>   8888R     8888R     888E  9888 'Y"  d888 '88%" 
 888R I888>   8888P     8888P     888E  9888      8888.+"    
u8888cJ888    *888>     *888>     888E  9888      8888L      
 "*888*P"     4888      4888      888&  ?8888u../ '8888c. .+ 
   'Y"        '888      '888      R888"  "8888P'   "88888%   
               88R       88R       ""      "P'       "YP'    
               88>       88>                                 
               48        48                                  
               '8        '8                                                          
"@

$BANNER_REINSTALL = @"
⚠️ REINSTALL OFFICE ???
( office is already installed )
"@

$BANNER_GET_IMAGE_PATH = @"
📥 PRESS [D] TO START DOWNLOADING IN BROWSER OR DO IT MANUALLY USING THE LINK:

$DOWNLOAD_URL

💿 THEN PRESS [ENTER] TO SELECT OFFICE IMAGE FILE
"@

$BANNER_PRODUCT_SELECTOR = @"
📦 CUSTOMIZE INSTALLATION
( select products to install )
"@

$BANNER_SETUP_OFFICE = @"
⚙️ SETUP OFFICE

DON'T OPEN APPS & DON'T CLOSE THIS WINDOW !!!
"@

$BANNER_SETUP_SUCCESS = @"
✅ SUCCESS
( office installed & activated )

PRESS ANY KEY TO EXIT...
"@

$BANNER_SETUP_ERROR = @"
❌ ERROR

PRESS ANY KEY TO EXIT...
"@