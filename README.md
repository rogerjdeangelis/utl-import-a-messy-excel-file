# utl-import-a-messy-excel-file
Import a messy excel file 
    Import a messy excel file                                                                                                 
                                                                                                                              
         Method                                                                                                               
                                                                                                                              
             a. modify registry typeguessrows (I use 32756)                                                                   
                May be different for your system (MS likes to change locations for no reason?)                                
                                                                                                                              
                regedit \HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\office\14.0\Access Conectivity Engine\Engines\Excel            
                                                                                                                              
             b. libname with testsize=32756;                                                                                  
                                                                                                                              
             c. proc contents _all_                                                                                           
                                                                                                                              
             d. Use passthru to EXCEL SQL query language to check long lengths                                                
                                                                                                                              
             e. Use libname to create SAS table                                                                               
                                                                                                                              
                                                                                                                              
    github                                                                                                                    
    https://github.com/rogerjdeangelis/utl-import-a-messy-excel-file                                                          
                                                                                                                              
    SAS forum                                                                                                                 
    https://tinyurl.com/qlvmdcp                                                                                               
    https://communities.sas.com/t5/SAS-Text-and-Content-Analytics/Text-Import-only-pulling-in-header-row/m-p/610042           
                                                                                                                              
    *_                   _                                                                                                    
    (_)_ __  _ __  _   _| |_                                                                                                  
    | | '_ \| '_ \| | | | __|                                                                                                 
    | | | | | |_) | |_| | |_                                                                                                  
    |_|_| |_| .__/ \__,_|\__|                                                                                                 
            |_|                                                                                                               
    ;                                                                                                                         
                                                                                                                              
    Download workbook from SAS  Forum                                                                                         
                                                                                                                              
    *            _               _                                                                                            
      ___  _   _| |_ _ __  _   _| |_                                                                                          
     / _ \| | | | __| '_ \| | | | __|                                                                                         
    | (_) | |_| | |_| |_) | |_| | |_                                                                                          
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                         
                    |_|                                                                                                       
    ;                                                                                                                         
                                                                                                                              
    libname xin "d:/xls/sampledata.xlsx" textsize=32000;                                                                      
                                                                                                                              
    data want;                                                                                                                
      set XIN.'out1$'n;                                                                                                       
    run;quit;                                                                                                                 
                                                                                                                              
    NOTE: There were 19953 observations read from the data set XIN.'out1$'n.                                                  
    NOTE: The data set WORK.WANT has 19953 observations and 25 variables.                                                     
    NOTE: DATA statement used (Total process time):                                                                           
          real time           8.41 seconds                                                                                    
                                                                                                                              
                                                                                                                              
    Middle Observation(9976 ) of Last dataset = WORK.WANT - Total Obs 19953                                                   
                                                                                                                              
                                             Value                                                                            
    Variable                        Type     Truncated(ob=9,976)  Passthru name                                               
    =========                       ===      ===================  ====================                                        
                                                                                                                              
    TEXT                             C166     RT @1075rosebud:    Text                                                        
    COUNTRY                          C27                          Country                                                     
    HASHTAGS                         C116     ClimateChange       Hashtags                                                    
    ID                               C46      tag:search.twitt    ID                                                          
    RETWEET                          C10                          Retweet                                                     
    LINKS                            C62      http://twitter.c    Links                                                       
    LOCATION_TYPE                    C7                           Location Type                                               
    LOCATION_COORDINATES             C146                         Location Coordinates                                        
    LOCATION_DISPLAY_NAME            C52                          Location Display Name                                       
    MEDIA_DISPLAY_URL                C110                         Media Display URL                                           
    MEDIA_URL                        C190                         Media URL                                                   
    REAL_NAME                        C20      ????                Real Name                                                   
    SOURCE                           C32      Echofon             Source                                                      
    TWEET_URL                        C518                         Tweet URL                                                   
    USER_BIO_SUMMARY                 C160     The Arab world c    User Bio Summary                                            
    USER_LOCATION                    C30      Melbourne Austra    User Location                                               
    USER_MENTION                     C126     Susan Steel         User Mention                                                
    USER_MENTION_USERNAME            C111     1075rosebud         user_mention_username                                       
    USER_TWITTER_PAGE                C38      http://www.twitt    user_twitter_page                                           
    USERNAME                         C15      khaladk             Username                                                    
                                                                                                                              
                                                                                                                              
     -- NUMERIC --                                                                                                            
    FAVORITES                        N8       19                  Favorites                                                   
    FOLLOWERS                        N8       696                 Followers                                                   
    FRIENDS                          N8       632                 Friends                                                     
    POSTED_TIME                      N8       20030.289375        Posted Time                                                 
    STATUSES_COUNT_                  N8       32291               Statuses Count:                                             
                                                                                                                              
    *                                                                                                                         
     _ __  _ __ ___   ___ ___  ___ ___                                                                                        
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                       
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                                       
    | .__/|_|  \___/ \___\___||___/___/                                                                                       
    |_|                                                                                                                       
    ;                                                                                                                         
                                                                                                                              
    Download the workbook from SAS forum                                                                                      
    *                                  _ _ _                                                                                  
      __ _     _ __ ___  __ _  ___  __| (_) |_                                                                                
     / _` |   | '__/ _ \/ _` |/ _ \/ _` | | __|                                                                               
    | (_| |_  | | |  __/ (_| |  __/ (_| | | |_                                                                                
     \__,_(_) |_|  \___|\__, |\___|\__,_|_|\__|                                                                               
                        |___/                                                                                                 
    ;                                                                                                                         
                                                                                                                              
    modify registry typeguessrows (I use 32756)                                                                               
    May be different for your system (MS likes to change locations for no reason?)                                            
                                                                                                                              
    Right mouse on typeguessrows and select modify                                                                            
    select decimal                                                                                                            
    enter 32756                                                                                                               
                                                                                                                              
    *_        _ _ _                                                                                                           
    | |__    | (_) |__  _ __   __ _ _ __ ___   ___                                                                            
    | '_ \   | | | '_ \| '_ \ / _` | '_ ` _ \ / _ \                                                                           
    | |_) |  | | | |_) | | | | (_| | | | | | |  __/                                                                           
    |_.__(_) |_|_|_.__/|_| |_|\__,_|_| |_| |_|\___|                                                                           
                                                                                                                              
    ;                                                                                                                         
                                                                                                                              
    libname xin "d:/xls/sampledata.xlsx" textsize=32000;                                                                      
                                                                                                                              
    *                         _             _                                                                                 
      ___      ___ ___  _ __ | |_ ___ _ __ | |_ ___                                                                           
     / __|    / __/ _ \| '_ \| __/ _ \ '_ \| __/ __|                                                                          
    | (__ _  | (_| (_) | | | | ||  __/ | | | |_\__ \                                                                          
     \___(_)  \___\___/|_| |_|\__\___|_| |_|\__|___/                                                                          
                                                                                                                              
    ;                                                                                                                         
                                                                                                                              
    proc contents data=xin._all_ position;                                                                                    
    run;quit;                                                                                                                 
                                                                                                                              
    Libref         XIN                                                                                                        
    Engine         EXCEL                                                                                                      
    Physical Name  d:/xls/sampledata.xlsx                                                                                     
                                                   DBMS                                                                       
                                   Member          Member                                                                     
    #  Name                        Type    Vars    Type                                                                       
                                                                                                                              
    1  out1$                       DATA     25     TABLE                                                                      
    2  out1$_xlnm#_FilterDatabase  DATA     25     TABLE  * use this for passthru (same as out$1)                             
                                                                                                                              
    Use the label with passthru. Bracket name in label has spaces                                                             
                                                                                                                              
    Variable                 Type    Len  Label                                                                               
                                                                                                                              
    TEXT                     Char    166  Text                                                                                
    COUNTRY                  Char     27  Country                                                                             
    FAVORITES                Num       8  Favorites                                                                           
    FOLLOWERS                Num       8  Followers                                                                           
    FRIENDS                  Num       8  Friends                                                                             
    HASHTAGS                 Char    116  Hashtags                                                                            
    ID                       Char     46  ID                                                                                  
    RETWEET                  Char     10  Retweet                                                                             
    LINKS                    Char     62  Links                                                                               
    LOCATION_TYPE            Char      7  Location Type                                                                       
    LOCATION_COORDINATES     Char    146  Location Coordinates                                                                
    LOCATION_DISPLAY_NAME    Char     52  Location Display Name                                                               
    MEDIA_DISPLAY_URL        Char    110  Media Display URL                                                                   
    MEDIA_URL                Char    190  Media URL                                                                           
    POSTED_TIME              Num       8  Posted Time                                                                         
    REAL_NAME                Char     20  Real Name                                                                           
    SOURCE                   Char     32  Source                                                                              
    STATUSES_COUNT_          Num       8  Statuses Count:                                                                     
    TWEET_URL                Char    518  Tweet URL                                                                           
    USER_BIO_SUMMARY         Char    160  User Bio Summary                                                                    
    USER_LOCATION            Char     30  User Location                                                                       
    USER_MENTION             Char    126  User Mention                                                                        
    USER_MENTION_USERNAME    Char    111  user_mention_username                                                               
    USER_TWITTER_PAGE        Char     38  user_twitter_page                                                                   
    USERNAME                 Char     15  Username                                                                            
                                                                                                                              
    *    _          _               _      _                  _   _                                                           
      __| |     ___| |__   ___  ___| | __ | | ___ _ __   __ _| |_| |__                                                        
     / _` |    / __| '_ \ / _ \/ __| |/ / | |/ _ \ '_ \ / _` | __| '_ \                                                       
    | (_| |_  | (__| | | |  __/ (__|   <  | |  __/ | | | (_| | |_| | | |                                                      
     \__,_(_)  \___|_| |_|\___|\___|_|\_\ |_|\___|_| |_|\__, |\__|_| |_|                                                      
                                                        |___/                                                                 
    ;                                                                                                                         
                                                                                                                              
    proc sql dquote=ansi;                                                                                                     
      connect to excel (Path="d:/xls/sampledata.xlsx");                                                                       
        create                                                                                                                
            table varlen as                                                                                                   
        select * from connection to Excel                                                                                     
            (                                                                                                                 
             Select                                                                                                           
                max(len(text)) as LTEXT                                                                                       
               ,max(len(Hashtags)) as LHASHTAGS                                                                               
               ,max(len([Tweet URL            ])) as LTweet_URL                                                               
               ,max(len([Media URL            ])) as LMedia_URL                                                               
               ,max(len([Media Display URL    ])) as LMedia_Display_URL                                                       
               ,max(len([User Bio Summary     ])) as LUser_Bio_Summary                                                        
               ,max(len([User Mention         ])) as LUser_Mention                                                            
               ,max(len([user_mention_username])) as Luser_mention_username                                                   
             from                                                                                                             
                [out1$_xlnm#_FilterDatabase]                                                                                  
            );                                                                                                                
        disconnect from Excel;                                                                                                
    quit;                                                                                                                     
                                                                                                                              
    proc transpose data=varlen out=varXpo;                                                                                    
    run;quit;                                                                                                                 
                                                                                                                              
                                Length                                                                                        
                                                                                                                              
    LTEXT                        166  chk                                                                                     
    LHASHTAGS                    116  chk                                                                                     
    LTWEET_URL                   518  chk (max sometimes 1024)                                                                
    LMEDIA_DISPLAY_URL           110  chk                                                                                     
    LMEDIA_URL                   190  chk                                                                                     
    LUSER_BIO_SUMMARY            160  chk                                                                                     
    LUSER_MENTION                126  chk                                                                                     
    LUSER_MENTION_USERNAME       111  chk                                                                                     
                                                                                                                              
    *    _                         _         _        _     _                                                                 
      __| |     ___ _ __ ___  __ _| |_ ___  | |_ __ _| |__ | | ___                                                            
     / _` |    / __| '__/ _ \/ _` | __/ _ \ | __/ _` | '_ \| |/ _ \                                                           
    | (_| |_  | (__| | |  __/ (_| | ||  __/ | || (_| | |_) | |  __/                                                           
     \__,_(_)  \___|_|  \___|\__,_|\__\___|  \__\__,_|_.__/|_|\___|                                                           
                                                                                                                              
    ;                                                                                                                         
                                                                                                                              
    data want;                                                                                                                
      set XIN.'out1$'n;                                                                                                       
    run;quit;                                                                                                                 
                                                                                                                              
    NOTE: There were 19953 observations read from the data set XIN.'out1$'n.                                                  
    NOTE: The data set WORK.WANT has 19953 observations and 25 variables.                                                     
    NOTE: DATA statement used (Total process time):                                                                           
          real time           8.41 seconds                                                                                    
                                                                                                                              
                                                                                                                              
    Middle Observation(9976 ) of Last dataset = WORK.WANT - Total Obs 19953                                                   
                                                                                                                              
                                              Value                                                                           
    Variable                        Type      Truncate(ob=9976)   Passthru name                                               
    =========                       ===       =================   ====================                                        
                                                                                                                              
    TEXT                             C166     RT @1075rosebud:    Text                                                        
    COUNTRY                          C27                          Country                                                     
    HASHTAGS                         C116     ClimateChange       Hashtags                                                    
    ID                               C46      tag:search.twitt    ID                                                          
    RETWEET                          C10                          Retweet                                                     
    LINKS                            C62      http://twitter.c    Links                                                       
    LOCATION_TYPE                    C7                           Location Type                                               
    LOCATION_COORDINATES             C146                         Location Coordinates                                        
    LOCATION_DISPLAY_NAME            C52                          Location Display Name                                       
    MEDIA_DISPLAY_URL                C110                         Media Display URL                                           
    MEDIA_URL                        C190                         Media URL                                                   
    REAL_NAME                        C20      ????                Real Name                                                   
    SOURCE                           C32      Echofon             Source                                                      
    TWEET_URL                        C518                         Tweet URL                                                   
    USER_BIO_SUMMARY                 C160     The Arab world c    User Bio Summary                                            
    USER_LOCATION                    C30      Melbourne Austra    User Location                                               
    USER_MENTION                     C126     Susan Steel         User Mention                                                
    USER_MENTION_USERNAME            C111     1075rosebud         user_mention_username                                       
    USER_TWITTER_PAGE                C38      http://www.twitt    user_twitter_page                                           
    USERNAME                         C15      khaladk             Username                                                    
                                                                                                                              
                                                                                                                              
     -- NUMERIC --                                                                                                            
    FAVORITES                        N8       19                  Favorites                                                   
    FOLLOWERS                        N8       696                 Followers                                                   
    FRIENDS                          N8       632                 Friends                                                     
    POSTED_TIME                      N8       20030.289375        Posted Time                                                 
    STATUSES_COUNT_                  N8       32291               Statuses Count:                                             
                                                                                                                              
                                                                                                                              
