
import { CellHyperlinkValue, Workbook } from 'exceljs'
import Validator from 'validator'
import { validateEmail, validatePassword } from './validationHelpers';
import { ImportSeedPayload } from './interface';

const GROUP_TYPE = ['SERVICE', 'TEAM', 'OFFICE', 'ORGANIZATION']
const ROLE = ['USER', 'MANAGER', 'ADMIN', 'SUPERADMIN', 'GUEST']
const POSITION = ['REGISTERED_CUSTOMS_REPRESENTATIVE', 'CUSTOMS_MANAGER', 'CUSTOMS_TEAM_MANAGER', 'INFORMATION_SYSTEMS_MANAGER', 'NABU_ADMINISTRATOR']
const LANGUAGE = ['fr', 'en']
const TRANSPORT_TYPE = ['FTP', 'SFTP', 'API']

function parseOrganizationWorksheet(workbook: Workbook, mapGroupNameToGroupId: Map<string, string>): any[] {
    const organizationWorksheet = workbook.getWorksheet('Organization & Groups')
    const groups: any[] = []
    organizationWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const groupName = row.getCell(1).value?.toString().trim()
            const groupType = row.getCell(2).value?.toString().trim()
            const parent = row.getCell(3).value?.toString().trim()

            try {
                if (!GROUP_TYPE.includes(groupType as string)) {
                    throw new Error("Invalid groupType")
                }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (parent && !mapGroupNameToGroupId.get(parent)) {
                    throw new Error("Invalid parent, parent must be defined before using it.")
                }

            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Organization & Groups)`)
            }

            mapGroupNameToGroupId.set(groupName, groupName)
            groups.push({
                groupName,
                groupType,
                parent
            })
        }
    });
    return groups
}

function parseUsersWorksheet(workbook: Workbook, mapGroupNameToGroupId: Map<string, string>): any[] {
    const usersWorksheet = workbook.getWorksheet('Users')
    const users: any[] = []
    usersWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const email = (row.getCell(1).value as CellHyperlinkValue).text.trim()
            const role = row.getCell(2).value?.toString().trim()
            const active = row.getCell(3).value?.toString().trim()
            const groupName = row.getCell(4).value?.toString().trim()
            const firstName = row.getCell(5).value?.toString().trim()
            const lastName = row.getCell(6).value?.toString().trim()
            const position = row.getCell(7).value?.toString().trim()
            const language = row.getCell(8).value?.toString().trim()
            const password = row.getCell(9).value?.toString().trim()

            try {
                if (!validateEmail(email)) {
                    throw new Error("Invalid email")
                }

                if (!ROLE.includes(role as string)) {
                    throw new Error("Invalid role")
                }

                if (!['oui', 'non'].includes(active as string)) {
                    throw new Error("Invalid value for active")
                }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("Group name is not defined in the previous worksheet")
                }

                if (!firstName) {
                    throw new Error("Invalid first name")
                }

                if (!lastName) {
                    throw new Error("Invalid last name")
                }

                if (!POSITION.includes(position as string)) {
                    throw new Error("Invalid position")
                }

                if (!LANGUAGE.includes(language as string)) {
                    throw new Error("Invalid language")
                }

                if (!validatePassword(password as string)) {
                    throw new Error("Invalid password format")
                }

            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Users)`)
            }

            users.push({
                email,
                role,
                active: active === "oui" ? true : false,
                groupName,
                firstName,
                lastName,
                position,
                password
            })
        }
    })
    return users
}

function parseCompanyRegistrationNumbersWorksheet(workbook: Workbook, mapGroupNameToGroupId: Map<string, string>): any[] {
    const companyRegistrationNumbersWorksheet = workbook.getWorksheet('Company Registration Numbers')
    const companyRegistrationNumbers: any[] = []
    companyRegistrationNumbersWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const value = row.getCell(1).value?.toString().trim()
            const country = row.getCell(2).value?.toString().trim()
            const groupName = row.getCell(3).value?.toString().trim()

            try {
                if (!value) {
                    throw new Error("Invalid value")
                }

                if (!country) {
                    throw new Error("Invalid country")
                }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("Group name is not defined in the previous worksheet")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Company Registration Numbers)`)
            }

            companyRegistrationNumbers.push({
                value,
                country,
                groupName
            })
        }
    })
    return companyRegistrationNumbers
}

function parseAddressesWorksheet(workbook: Workbook, mapGroupNameToGroupId: Map<string, string>): any[] {
    const addressWorksheet = workbook.getWorksheet('Addresses')
    const addresses: any[] = []
    addressWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const givenName = row.getCell(1).value?.toString().trim()
            const surname = row.getCell(2).value?.toString().trim()
            const companyName = row.getCell(3).value?.toString().trim()
            const streetNumber = row.getCell(4).value?.toString().trim()
            const streetName = row.getCell(5).value?.toString().trim()
            const floor = row.getCell(6).value?.toString().trim()
            const town = row.getCell(7).value?.toString().trim()
            const region = row.getCell(8).value?.toString().trim()
            const postCode = row.getCell(9).value?.toString().trim()
            const country = row.getCell(10).value?.toString().trim()
            const tag = row.getCell(11).value?.toString().trim()
            const groupName = row.getCell(12).value?.toString().trim()

            try {
                if (!givenName) {
                    throw new Error("Invalid givenName")
                }

                if (!surname) {
                    throw new Error("Invalid surname")
                }

                if (!companyName) {
                    throw new Error("Invalid companyName")
                }

                if (!streetNumber) {
                    throw new Error("Invalid streetNumber")
                }

                if (!streetName) {
                    throw new Error("Invalid streetName")
                }

                if (!floor) {
                    throw new Error("Invalid floor")
                }

                if (!town) {
                    throw new Error("Invalid town")
                }

                if (!region) {
                    throw new Error("Invalid region")
                }

                if (!postCode) {
                    throw new Error("Invalid postCode")
                }

                if (!country) {
                    throw new Error("Invalid country")
                }

                if (!tag) {
                    throw new Error("Invalid tag")
                }

                if (!groupName) {
                    throw new Error("Invalid group name")
                }

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("Group name is not defined in the previous worksheet")
                }

            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Addresses)`)
            }

            addresses.push({
                givenName,
                surname,
                companyName,
                streetNumber,
                streetName,
                floor,
                town,
                region,
                postCode,
                country,
                tag,
                groupName
            })
        }
    })
    return addresses
}

function parseSchemasWorksheet(workbook: Workbook, mapSchemaNameToSchemaId: Map<string, string>): any[] {
    const schemasWorksheet = workbook.getWorksheet("Schemas")
    const schemas: any[] = []
    schemasWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const name = row.getCell(1).value?.toString().trim()
            let content = row.getCell(2).value?.toString().trim()

            try {
                if (!name) {
                    throw new Error("Invalid name")
                }

                if (!content) {
                    throw new Error("Invalid content")
                }

                try {
                    content = JSON.parse(content)
                } catch (_) {
                    throw new Error("Content is an invalid JSON")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Schemas)`)
            }

            mapSchemaNameToSchemaId.set(name, name)
            schemas.push({
                name,
                content
            })
        }
    })
    return schemas;
}

function parseExportDataTemplatesWorksheet(workbook: Workbook, mapSchemaNameToSchemaId: Map<string, string>, mapExportDataTemplateNameToExportDataTemplateId: Map<string, string>): any[] {
    const exportDataTemplatesWorksheet = workbook.getWorksheet("Export Data Templates")
    const exportDataTemplates: any[] = []
    exportDataTemplatesWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const name = row.getCell(1).value?.toString().trim()
            const schemaName = row.getCell(2).value?.toString().trim()
            const isNew = row.getCell(3).value?.toString().trim()

            try {
                if (!name) {
                    throw new Error("Invalid name")
                }

                if (isNew) {
                    if (!['oui', 'non'].includes(isNew)) {
                        throw new Error("Invalid New value")
                    }
                    if (isNew === "non" && Validator.isUUID(schemaName as string)) {
                        throw new Error(`If new is false then the schema name should be an valid uuid received: ${schemaName}`)
                    }
                }

                if (!mapSchemaNameToSchemaId.get(schemaName as string)) {
                    throw new Error("Schema name is not defined in the previous worksheet")
                }


            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Export Data Templates)`)
            }

            mapExportDataTemplateNameToExportDataTemplateId.set(name, name)

            exportDataTemplates.push({
                name,
                schemaName,
                isNew: isNew === "oui" ? true : false
            })
        }
    })
    return exportDataTemplates
}

function parseTransportsWorksheet(workbook: Workbook, mapTransportNameToTransportId: Map<string, string>) {
    const transportsWorksheet = workbook.getWorksheet("Transports")
    const transports: any[] = []
    transportsWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const type = row.getCell(1).value?.toString().trim()
            const name = row.getCell(2).value?.toString().trim()
            const address = row.getCell(3).value?.toString().trim()
            const port = row.getCell(4).value?.toString().trim()
            const username = (row.getCell(5).value as CellHyperlinkValue).text.toString().trim()
            const password = row.getCell(6).value?.toString().trim()

            try {
                if (!TRANSPORT_TYPE.includes(type as string)) {
                    throw new Error("Invalid transport type")
                }

                if (!name) {
                    throw new Error("Invalid name")
                }

                if (!address) {
                    throw new Error("Invalid address")
                }

                if (!port) {
                    throw new Error("Invalid port")
                }

                if (!validateEmail(username)) {
                    throw new Error("Invalid username")
                }

                if (!validatePassword(password as string)) {
                    throw new Error("Invalid password format")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Transports)`)
            }

            mapTransportNameToTransportId.set(name, name)

            transports.push({
                type,
                name,
                address,
                port,
                username,
                password
            })
        }
    })
    return transports
}

function parseConfigurationsWorksheet(workbook: Workbook, mapExportDataTemplateNameToExportDataTemplateId: Map<string, string>, mapTransportNameToTransportId: Map<string, string>, mapGroupNameToGroupId: Map<string, string>): any[] {
    const configurationWorksheet = workbook.getWorksheet("Configurations")
    const configurations: any[] = []
    configurationWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const name = row.getCell(1).value?.toString().trim()
            const exportDataTemplateName = row.getCell(2).value?.toString().trim()
            const transportName = row.getCell(3).value?.toString().trim()
            const groupName = row.getCell(4).value?.toString().trim()
            const preset = row.getCell(5).value?.toString().trim()

            try {
                if (!name) {
                    throw new Error("Invalid name")
                }

                if (!mapExportDataTemplateNameToExportDataTemplateId.get(exportDataTemplateName as string)) {
                    throw new Error("ExportDataTemplateName is not defined in the previous worksheet")
                }

                if (!mapTransportNameToTransportId.get(transportName as string)) {
                    throw new Error("TransportName is not defined in the previous worksheet")
                }

                if (!mapGroupNameToGroupId.get(groupName as string)) {
                    throw new Error("GroupName is not defined in the previous worksheet")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Configurations)`)
            }

            configurations.push({
                name,
                exportDataTemplateName,
                transportName,
                groupName,
                preset
            })
        }
    })
    return configurations
}

function parseSettingsWorksheet(workbook: Workbook, mapGroupNameToGroupId: Map<string, string>): any[] {
    const settingsWorksheet = workbook.getWorksheet("Settings")
    const settings: any[] = []
    settingsWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const key = row.getCell(1).value?.toString().trim()
            const value = row.getCell(2).value?.toString().trim()
            const groupName = row.getCell(3).value?.toString().trim()

            try {
                if (!key) {
                    throw new Error("Invalid key")
                }

                if (!value) {
                    throw new Error("Invalid value")
                }

                if (!mapGroupNameToGroupId.get(groupName as string)) {
                    throw new Error("GroupName is not defined in the previous worksheet")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Settings)`)
            }

            settings.push({
                key,
                value,
                groupName
            })
        }
    })
    return settings
}


export default function processWorkbook(workbook: any) {
    /**
     * Parse the excel file
     */
    const jsonData: ImportSeedPayload = {
        organizations: [],
        groups: [],
        users: [],
        companyRegistrationNumbers: [],
        exportDataTemplates: [],
        addresses: [],
        schemas: [],
        transports: [],
        configurations: [],
        settings: [],
    };

    const mapGroupNameToGroupId = new Map()
    const mapSchemaNameToSchemaId = new Map()
    const mapExportDataTemplateNameToExportDataTemplateId = new Map()
    const mapTransportNameToTransportId = new Map()

    jsonData.groups = parseOrganizationWorksheet(workbook, mapGroupNameToGroupId)
    const hasAnOrganization = jsonData.groups.some(group => group.groupType == "ORGANIZATION")

    if (!hasAnOrganization) {
        throw new Error(`At least a group should be of type of ${"ORGANIZATION"}`)
    }

    jsonData.users = parseUsersWorksheet(workbook, mapGroupNameToGroupId)

    jsonData.companyRegistrationNumbers = parseCompanyRegistrationNumbersWorksheet(workbook, mapGroupNameToGroupId)

    jsonData.addresses = parseAddressesWorksheet(workbook, mapGroupNameToGroupId)

    jsonData.schemas = parseSchemasWorksheet(workbook, mapSchemaNameToSchemaId)

    jsonData.companyRegistrationNumbers = parseExportDataTemplatesWorksheet(workbook, mapSchemaNameToSchemaId, mapExportDataTemplateNameToExportDataTemplateId)

    jsonData.transports = parseTransportsWorksheet(workbook, mapTransportNameToTransportId)

    jsonData.configurations = parseConfigurationsWorksheet(workbook, mapExportDataTemplateNameToExportDataTemplateId, mapTransportNameToTransportId, mapGroupNameToGroupId)

    jsonData.settings = parseSettingsWorksheet(workbook, mapGroupNameToGroupId)


    return jsonData
}
