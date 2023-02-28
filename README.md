## SYNOPSIS
    Script to retrieve data from Nexthink via NXQL API

## DESCRIPTION
    GUI PowerShell script to to retrieve data from Nexthink via NXQL API.
    Create to use on multi engine Nexthink Experience environments.
    The result file will contains merged output from all connected engines,
    without any additional headers and blank lines.

## INPUTS
    Portal FQDN
    Username
    Password
    NXQL Query

## OUTPUTS
    Merged Nexthink engines output

## NOTES
    Version:            1.04
    Author:             Stanislaw Horna
    Mail:               stanislaw.horna@atos.net
    Creation Date:      16-Feb-2023
    ChangeLog:

    Date            Who                     What
    2023-02-18      Stanislaw Horna         Basic query validation added;
                                            Mechanism to keep credentials if valid;
                                            MoJ select environment GUI form added.

    2023-02-24      Stanislaw Horna         Show password button added;
                                            Connecting to portal with ENTER key;
                                            More accurate error handling;
                                            Possibility to change user, after establishing connection.
    
    2023-02-27      Stanislaw Horna         Better Handling if unable to create neccesary variables;
                                            Handling for running on unsupported environment;
                                            Handling for no write permission;
                                            Handling for running multiple apps at the same time.

    2023-02-28      Stanislaw Horna         Error Handling for invalid query - 
                                            returns the same message as NXQL WebEditor.