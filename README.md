# paloexport

Reinventing the wheel, one Palo XML schema at a time

pan_security_to_excel_xml - converts Panorama XML device-group security rules to excel
Usage: use below flags
  - '-f', '--file', required=True, help='XML filename'
  - '-d', '--device', required=True, help='device entry name - default is localhost.localdomain for a Panorama'
  - '-g', '--devicegroup', required=True, help='device group name'
  - '-r', '--rulebase', required=True, help='rulebase e.g. pre-rulebase'
