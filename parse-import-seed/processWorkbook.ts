
import ExcelJS, { CellHyperlinkValue } from 'exceljs'
import * as types from "./types";
import { validateEmail, validatePassword } from './validationHelpers'; // ok
import Validator from 'validator' // library


export function processWorkbook(workbook: any) {

    // Since we validate every sheet and columns, we can now access it safely.
    const organizationWorksheet = workbook.getWorksheet('Organization & Groups')
    const mapGroupNameToGroupId = new Map()
    let groupsData: any[] = []

    organizationWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const groupName = row.getCell(1).value?.toString().trim()
            const groupType = row.getCell(2).value?.toString().trim()
            const parent = row.getCell(3).value?.toString().trim()

            try {
                // TODO - fix
                // if (!Object.values(api.Type).includes(groupType as api.Type)) {
                //     throw new Error("Invalid groupType")
                // }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (parent && !mapGroupNameToGroupId.get(parent)) {
                    throw new Error("Invalid parent, parent must be defined before using it.")
                }

            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Organization & Groups)`)
            }

            mapGroupNameToGroupId.set(groupName, groupName)
            groupsData.push({
                groupName,
                groupType,
                parent
            })
        }
    });

    // TODO - fix
    // const hasAnOrganization = groupsData.some(group => group.groupType == types.Type.Organization)

    // if (!hasAnOrganization) {
    //     throw new Error(`At least a group should be of type of ${types.Type.Organization}`)
    // }

    const usersWorksheet = workbook.getWorksheet('Users')
    const usersData: any[] = []

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
                if (!validateEmail(email as string)) {
                    throw new Error("Invalid email")
                }

                // TODO - fix
                // if (!Object.values(api.Role).includes(role as api.Role)) {
                //     throw new Error("Invalid role")
                // }

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

                // TODO - fix
                // if (!Object.values(api.Position).includes(position as api.Position)) {
                //     throw new Error("Invalid position")
                // }

                // if (!Object.values(api.Language).includes(language as api.Language)) {
                //     throw new Error("Invalid language")
                // }

                if (!validatePassword(password as string)) {
                    throw new Error("Invalid password format")
                }

            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Users)`)
            }

            usersData.push({
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


    const companyRegistrationNumbersWorksheet = workbook.getWorksheet('Company Registration Numbers')
    const companyRegistrationsData: any[] = []
    companyRegistrationNumbersWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const value = row.getCell(1).value?.toString().trim()
            const country = row.getCell(2).value?.toString().trim()
            const groupName = row.getCell(3).value?.toString().trim()

            try {
                if (!value) {
                    throw new Error("Invalid value")
                }

                if (!Validator.isISO31661Alpha3(country as string)) {
                    throw new Error("Invalid country")
                }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("Group name is not defined in the previous worksheet")
                }
            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Company Registration Numbers)`)
            }

            companyRegistrationsData.push({
                value,
                country,
                groupName
            })
        }
    })


    const addressWorksheet = workbook.getWorksheet('Addresses')
    const addressesData: any[] = []

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

                if (!Validator.isISO31661Alpha3(country as string)) {
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

            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Addresses)`)
            }

            addressesData.push({
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

    const schemasWorksheet = workbook.getWorksheet("Schemas")
    const mapSchemaNameToSchemaId = new Map()
    const schemasData: any[] = []

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
            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Schemas)`)
            }

            mapSchemaNameToSchemaId.set(name, name)
            schemasData.push({
                name,
                content
            })
        }
    })


    const exportDataTemplatesWorksheet = workbook.getWorksheet("Export Data Templates")
    const mapExportDataTemplateNameToExportDataTemplateId = new Map()
    const exportDataTemplatesData: any[] = []

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

                if (!mapSchemaNameToSchemaId.get(schemaName)) {
                    throw new Error("Schema name is not defined in the previous worksheet")
                }


            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Export Data Templates)`)
            }

            mapExportDataTemplateNameToExportDataTemplateId.set(name, name)

            exportDataTemplatesData.push({
                name,
                schemaName,
                isNew: isNew === "oui" ? true : false
            })
        }
    })

    const transportsWorksheet = workbook.getWorksheet("Transports")
    const mapTransportNameToTransportId = new Map()
    const transportsData: any[] = []

    transportsWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const type = row.getCell(1).value?.toString().trim()
            const name = row.getCell(2).value?.toString().trim()
            const address = row.getCell(3).value?.toString().trim()
            const port = row.getCell(4).value?.toString().trim()
            const username = (row.getCell(5).value as CellHyperlinkValue).text.toString().trim()
            const password = row.getCell(6).value?.toString().trim()

            try {
                // TODO -fix
                // if (!Object.values(api.Type2).includes(type as api.Type2)) {
                //     throw new Error("Invalid transport type")
                // }

                if (!name) {
                    throw new Error("Invalid name")
                }

                if (!address) {
                    throw new Error("Invalid address")
                }

                if (!port) {
                    throw new Error("Invalid port")
                }

                if (!validateEmail(username as string)) {
                    throw new Error("Invalid username")
                }

                if (!validatePassword(password as string)) {
                    throw new Error("Invalid password format")
                }
            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Transports)`)
            }

            mapTransportNameToTransportId.set(name, name)

            transportsData.push({
                type,
                name,
                address,
                port,
                username,
                password
            })
        }
    })

    const configurationWorksheet = workbook.getWorksheet("Configurations")
    const configurationsData: any[] = []

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

                if (!mapExportDataTemplateNameToExportDataTemplateId.get(exportDataTemplateName)) {
                    throw new Error("ExportDataTemplateName is not defined in the previous worksheet")
                }

                if (!mapTransportNameToTransportId.get(transportName)) {
                    throw new Error("TransportName is not defined in the previous worksheet")
                }

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("GroupName is not defined in the previous worksheet")
                }
            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Configurations)`)
            }

            configurationsData.push({
                name,
                exportDataTemplateName,
                transportName,
                groupName,
                preset
            })
        }
    })

    const settingsWorksheet = workbook.getWorksheet("Settings")
    const settingsData: any[] = []

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

                if (!mapGroupNameToGroupId.get(groupName)) {
                    throw new Error("GroupName is not defined in the previous worksheet")
                }
            } catch (err: any) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Settings)`)
            }

            settingsData.push({
                key,
                value,
                groupName
            })
        }
    })

    return {
      settingsData,
      configurationsData,
      transportsData,
      exportDataTemplatesData,
      schemasData,
      addressesData,
      companyRegistrationsData,
      usersData,
      groupsData
    }
}