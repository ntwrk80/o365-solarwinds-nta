import json
import urllib.request
import uuid
import os
import ipaddress
import json2xml


#Original starting code from Microsoft article:
#(https://docs.microsoft.com/en-us/office365/enterprise/office-365-ip-web-service)
#

# helper to call the webservice and parse the response
def webApiGet(methodName, instanceName, clientRequestId):
    ws = "https://endpoints.office.com"
    requestPath = ws + '/' + methodName + '/' + instanceName + '?clientRequestId=' + clientRequestId
    request = urllib.request.Request(requestPath)
    with urllib.request.urlopen(request) as response:
        return json.loads(response.read().decode())

def printXML(endpointSets):
    with open('importO365NTA.xml', 'w') as output:

        for endpointSet in endpointSets:
            if endpointSet['category'] in ('Optimize', 'Allow'):
                ips = endpointSet['ips'] if 'ips' in endpointSet else []
                category = endpointSet['category']
                serviceArea = endpointSet['serviceArea']
                # IPv4 strings have dots while IPv6 strings have colons
                ip4s = [ip for ip in ips if '.' in ip]
                tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
                udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
                flatIps.extend([(serviceArea, category, ip, tcpPorts, udpPorts) for ip in ip4s])
        print('IPv4 Firewall IP Address Ranges')
        #print (flatIps)
        currentServiceArea = " "
        output.write ("<AddressGroups xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns=\"http://tempuri.org/IPAddressGroupsSchema.xsd\">\n")
        for ip in flatIps:
            serviceArea = ip [0]
            if serviceArea != currentServiceArea:
                if currentServiceArea != " ":
                    output.write ("     </AddressGroup>\n")
                currentServiceArea = serviceArea
                output.write (f"     <AddressGroup enabled=\"true\" description=\"Office 365 {serviceArea}\">\n")
            ipNet = ipaddress.ip_network(ip[2])
            ipStart = ipNet[0]
            ipEnd = ipNet[-1]
            output.write (f"          <Range from=\"{ipStart}\" to=\"{ipEnd}\"/>\n")
        output.write ("     </AddressGroup>\n")
        output.write ("</AddressGroups>\n")
        #print('\n'.join(sorted(set([ip for (category, ip, tcpPorts, udpPorts) in flatIps]))))
        #print('URLs for Proxy Server')
        #print(','.join(sorted(set([url for (category, url, tcpPorts, udpPorts) in flatUrls]))))

        # TODO send mail (e.g. with smtplib/email modules) with new endpoints data


def main (argv):

    # path where client ID and latest version number will be stored
    datapath = 'endpoints_clientid_latestversion.txt'
    # fetch client ID and version if data exists; otherwise create new file
    if os.path.exists(datapath):
        with open(datapath, 'r') as fin:
            clientRequestId = fin.readline().strip()
            latestVersion = fin.readline().strip()
    else:
        clientRequestId = str(uuid.uuid4())
        latestVersion = '0000000000'
        with open(datapath, 'w') as fout:
            fout.write(clientRequestId + '\n' + latestVersion)
    version = webApiGet('version', 'Worldwide', clientRequestId)
    if version['latest'] > latestVersion:
        print('New version of Office 365 worldwide commercial service instance endpoints detected')
        # write the new version number to the data file
        with open(datapath, 'w') as fout:
            fout.write(clientRequestId + '\n' + version['latest'])
        # invoke endpoints method to get the new data
        endpointSets = webApiGet('endpoints', 'Worldwide', clientRequestId)
        # filter results for Allow and Optimize endpoints, and transform these into tuples with port and category
        flatUrls = []
        for endpointSet in endpointSets:
            if endpointSet['category'] in ('Optimize', 'Allow'):
                category = endpointSet['category']
                urls = endpointSet['urls'] if 'urls' in endpointSet else []
                tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
                udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
                flatUrls.extend([(category, url, tcpPorts, udpPorts) for url in urls])
        flatIps = []
        printXML(endpointSets)
    else:
        print('Office 365 worldwide commercial service instance endpoints are up-to-date')

if __name__ == '__main__':
    main(sys.argv)
