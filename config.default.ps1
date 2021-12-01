$config = @{
    'toggl_api_uri'                   = 'https://api.track.toggl.com/api/v8'
    'toggl_reports_api_uri'           = 'https://api.track.toggl.com/reports/api/v2'
    'api_token'                       = ""

    'client_name'                     = "Akvelon"
    'irregular_time_start'            = "21"
    'irregular_time_end'              = "8"
    'irregular_time_prefix'           = "Irregular hours - "
    'projects_without_irregular_time' = @( "Internal" )

    'effort_fractional_digits'        = 1

    'project_prefix'                  = @{
        'Internal'     = 'Internal  Project'
    }

    'excel_template_path' = "$PSScriptRoot\template.xlsx"
}