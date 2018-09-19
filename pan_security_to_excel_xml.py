#!/usr/bin/env python
import requests # Method of getting the XML information from PAN.
import xml.etree.ElementTree as ET # Difficult to use but good XML parser.
import xlsxwriter # Creates an Excel Spreadsheet.
import argparse

#ASSUMES CONFIG FILE IS running-config.xml


class Spreadsheet(object):
    """Create a spreadsheet from the XML document."""
    def __init__(self):
        self.name = None
        self.negate = "no"
        self.from_member = ""
        self.to_member = ""
        self.source = ""
        self.destination = ""
        self.sourceuser = ""
        self.category = ""
        self.application = ""
        self.service = ""
        self.hipprofile = ""
        self.action = None
        self.logend = ""
        self.logsetting = ""
        self.profile = ""
        self.description = None
        self.disabled = "no" # Set to no since the PAN might return nothing for permit.
        self.tag = ""
        self.expiration = None

    def writeRowHeaders(self):
        """Write the header row of the spreadsheet."""
        titles = ["Rule Name", "Negate", "From Zone", "To Zone", "Source", "Destination", "Source-User", "Category", "Application", "Service", "Hip-Profiles", "Action", "Log-End", "Log-Setting", "Profile", "Description", "Disabled", "Tag", "Expiration"]
        i = 0
        for title in titles:
            worksheet.write(0, i, title, bold)
            i += 1

    def setName(self, name):
        self.name = name

    def setNegate(self, negate):
        self.negate = negate

    def setFromMember(self, from_member):
        if not self.from_member == "": # If there are multiple entries add a comma to separate.
            self.from_member += chr(10)
        self.from_member +=str(from_member) # Concatenate each entry.

    def setToMember(self, to_member):
        if not self.to_member == "": # If there are multiple entries add a comma to separate.
            self.to_member += chr(10)
        self.to_member +=str(to_member) # Concatenate each entry.

    def setSource(self, source):
        if not self.source == "": # If there are multiple entries add a comma to separate.
            self.source += chr(10)
        self.source +=str(source) # Concatenate each entry.

    def setDestination(self, destination):
        if not self.destination == "": # If there are multiple entries add a comma to separate.
            self.destination += chr(10)
        self.destination +=str(destination) # Concatenate each entry.

    def setSourceUser(self, sourceuser):
        if not self.sourceuser == "": # If there are multiple entries add a comma to separate.
            self.sourceuser += chr(10)
        self.sourceuser +=str(sourceuser) # Concatenate each entry.

    def setCategory(self, category):
        if not self.category == "": # If there are multiple entries add a comma to separate.
            self.category += chr(10)
        self.category +=str(category) # Concatenate each entry.

    def setApplication(self, application):
        if not self.application == "": # If there are multiple entries add a comma to separate.
            self.application += chr(10)
        self.application +=str(application) # Concatenate each entry.

    def setService(self, service):
        if not self.service == "": # If there are multiple entries add a comma to separate.
            self.service += chr(10)
        self.service +=str(service) # Concatenate each entry.

    def setHipProfile(self, hipprofile):
        if not self.hipprofile == "": # If there are multiple entries add a comma to separate.
            self.hipprofile += chr(10)
        self.hipprofile +=str(hipprofile) # Concatenate each entry.

    def setAction(self, action):
        self.action = action

    def setLogEnd(self, logend):
        self.logend = logend

    def setLogSetting(self, logsetting):
        self.logsetting = logsetting

    def setProfile(self, profile):
        if not self.profile == "": # If there are multiple entries add a comma to separate.
            self.profile += chr(10)
        self.profile +=str(profile) # Concatenate each entry.

    def setDescription(self, description):
        self.description = description

    def setDisabled(self, disabled):
        self.disabled = disabled

    def setTag(self, tag):
        if not self.tag == "": # If there are multiple entries add a comma to separate.
            self.tag += chr(10)
        self.tag +=str(tag) # Concatenate each entry.

    def setExpiration(self, expiration):
        self.expiration = expiration

    def writeRow(self, row):
        """Writes row to Excel workbook"""
        # Insert validation later
        worksheet.write(row, 0, self.name, dataformat)
        worksheet.write(row, 1, self.negate, dataformat)
        worksheet.write(row, 2, self.from_member, dataformat)
        worksheet.write(row, 3, self.to_member, dataformat)
        worksheet.write(row, 4, self.source, dataformat)
        worksheet.write(row, 5, self.destination, dataformat)
        worksheet.write(row, 6, self.sourceuser, dataformat)
        worksheet.write(row, 7, self.category, dataformat)
        worksheet.write(row, 8, self.application, dataformat)
        worksheet.write(row, 9, self.service, dataformat)
        worksheet.write(row, 10, self.hipprofile, dataformat)
        worksheet.write(row, 11, self.action, dataformat)
        worksheet.write(row, 12, self.logend, dataformat)
        worksheet.write(row, 13, self.logsetting, dataformat)
        worksheet.write(row, 14, self.profile, dataformat)
        worksheet.write(row, 15, self.description, dataformat)
        worksheet.write(row, 16, self.disabled, dataformat)
        worksheet.write(row, 17, self.tag, dataformat)
        worksheet.write(row, 18, self.expiration, dataformat)

        print "Name: ", self.name
        print "Negate: ", self.negate
        print "From Zone: ", self.from_member
        print "To Zone: ", self.to_member
        print "Source: ", self.source
        print "Destination: ", self.destination
        print "Source-User: ", self.sourceuser
        print "Category: ", self.sourceuser
        print "Application: ", self.application
        print "Service: ", self.service
        print "Hip-Profiles: ", self.hipprofile
        print "Action: ", self.action
        print "Log-End: ", self.logend
        print "Log-Setting: ", self.logsetting
        print "Profile: ", self.profile
        print "Disabled: ", self.disabled
        print "Tag: ", self.tag
        print "Description: ", self.description
        print "Expiration: ", self.expiration
        print "\n"

    def newRow(self):
        """Prepares for new row by clearing variables in class"""
        excelobj.__init__()

def commandlineparser():
     global args
     parser = argparse.ArgumentParser(description='Convert Palo Alto Networks Panorama Device-Group Security rules from XML to Microsoft Excel.')
     parser.add_argument('-f', '--file', required=True, help='XML filename')
     parser.add_argument('-d', '--device', required=True, help='device entry name - default is localhost.localdomain for a Panorama')
     parser.add_argument('-g', '--devicegroup', required=True, help='device group name')
     parser.add_argument('-r', '--rulebase', required=True, help='rulebase e.g. pre-rulebase')
     args = parser.parse_args()


if __name__ == '__main__':

    #Get command line arguments
    commandlineparser()

    row = 0 # Used to track which excel row we are on while parsing XML.

    #url = "%s/api/?type=config&action=get&xpath=/config/devices/entry[@name=\'localhost.localdomain\']/device-group/entry[@name=\'%s\']/pre-rulebase/security/rules&key=%s" % (args.panorama, args.firewall, args.apikey)

    xml = ET.parse(args.file)
    document = xml.getroot()

    workbook = xlsxwriter.Workbook(args.device + '.' + args.devicegroup + "." + "security.policies.xlsx") # Create Excel spreadsheet.
    worksheet = workbook.add_worksheet() # Create new worksheet within the spreadsheet.

    bold = workbook.add_format({'bold': True}) # Cell formatting for row header

    dataformat = workbook.add_format() # Cell Formatting for data.
    dataformat.set_align('top')
    dataformat.set_text_wrap()
    worksheet.set_column(0,15, 20)

    excelobj = Spreadsheet()
    excelobj.writeRowHeaders() # Create friendly row headers in the spreadsheet.

    for root in document.findall("devices"): # Start after root (config)
        xpath_device = "entry[@name='%s']" % args.device
        for device in root.iterfind(xpath_device):
#            for device_group in device.iterfind(xpath_devicegroup):
            for top_level_element in device.findall("device-group"):
                xpath_devicegroup = "entry[@name='%s']" % args.devicegroup
                for device_group in top_level_element.iterfind(xpath_devicegroup):
                    for rulebase in device_group.findall(args.rulebase):
                        for security in rulebase.findall("security"):
                            for rules in security.findall("rules"):
                                for entries in rules:
                                    row += 1
                                    excelobj.setName(name=entries.attrib.get("name")) # Populate the rule description. Used attrib.get since name is a value within the tag.

                                    for negate in entries.findall("negate"):
                                        excelobj.setNegate(negate.text)

                                    for fromzone in entries.findall("from"): # From zone block
                                        for members in fromzone.findall("member"): # From zone block - members block
                                            excelobj.setFromMember(members.text)

                                    for tozone in entries.findall("to"): # To zone block
                                        for members in tozone.findall("member"): # To zone block - members block
                                            excelobj.setToMember(members.text)

                                    for source in entries.findall("source"): # From source block
                                        for members in source.findall("member"): # From source block - members block
                                            excelobj.setSource(members.text)

                                    for destination in entries.findall("destination"): # application block
                                        for members in destination.findall("member"): # application block - members block
                                            excelobj.setDestination(members.text)

                                    for sourceuser in entries.findall("source-user"): # application block
                                        for members in sourceuser.findall("member"): # application block - members block
                                            excelobj.setSourceUser(members.text)

                                    for category in entries.findall("category"): # application block
                                        for members in category.findall("member"): # application block - members block
                                            excelobj.setCategory(members.text)

                                    for application in entries.findall("application"): # application block
                                        for members in application.findall("member"): # application block - members block
                                            excelobj.setApplication(members.text)

                                    for service in entries.findall("service"): # application block
                                        for members in service.findall("member"): # application block - members block
                                            excelobj.setService(members.text)

                                    for hipprofile in entries.findall("hip-profiles"): # application block
                                        for members in hipprofile.findall("member"): # application block - members block
                                            excelobj.setHipProfile(members.text)

                                    for action in entries.findall("action"):
                                        excelobj.setAction(action.text)

                                    for logend in entries.findall("log-end"):
                                        excelobj.setLogEnd(logend.text)

                                    for logsetting in entries.findall("log-setting"):
                                        excelobj.setLogSetting(logsetting.text)

                                    for profset in entries.findall("profile-setting"):
                                        for prof in profset.findall("profiles"):
                                            for virus in prof.findall("virus"):
                                                for members in virus.findall("member"):
                                                    excelobj.setProfile(virus.tag + " - " + members.text)
                                            for spyware in prof.findall("spyware"):
                                                for members in spyware.findall("member"):
                                                    excelobj.setProfile(spyware.tag + " - " + members.text)
                                            for vuln in prof.findall("vulnerability"):
                                                for members in vuln.findall("member"):
                                                    excelobj.setProfile(vuln.tag + " - " + members.text)
                                            for url in prof.findall("url-filtering"):
                                                for members in url.findall("member"):
                                                    excelobj.setProfile(url.tag + " - " + members.text)
                                            for file in prof.findall("file-blocking"):
                                                for members in file.findall("member"):
                                                    excelobj.setProfile(file.tag + " - " + members.text)
                                            for group in prof.findall("group"):
                                                for members in file.findall("member"):
                                                    excelobj.setProfile(group.tag + " - " + members.text)

                                    for description in entries.findall("description"):
                                        excelobj.setDescription(description.text)

                                    for disabled in entries.findall("disabled"):
                                        excelobj.setDisabled(disabled.text)

                                    for tag in entries.findall("tag"): # application block
                                        for members in tag.findall("member"): # application block - members block
                                            excelobj.setTag(members.text)

                                    for expiration in entries.findall("schedule"):
                                        excelobj.setExpiration(expiration.text)

                                    excelobj.writeRow(row) # Write each row to the spreadsheet.
                                    excelobj.newRow() # Clear old values and start new row.


    workbook.close() # Close the spreadsheet since we are done with it now.
#
# # XML document structure
# # <repsonse>
# #   <result>
# #       <rules>
# #           <entry>
# #               <from>
# #                   <member>from zone</member>
# #               </from>
# #                <to>
# #                   <member>to zone</member>
# #                </to>
# #               <source>
# #                   <member>source network</member>
# #               </source>
# #               <destination>
# #                   <member>destination network</member>
# #               </destination>
# #               <application>
# #                   <member>application</member>
# #               </application>
# #               <action>
# #                   value
# #               </action>
# #               <description>
# #                   value
# #               </description>
# #               <disabled>
# #                   value
# #               </disabled>
# #           </entry>
# #       </rules>
# #   </result>
# # </repsonse>
