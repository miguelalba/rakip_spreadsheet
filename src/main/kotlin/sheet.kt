import com.gmail.gcolaianni5.jris.bean.Record
import com.gmail.gcolaianni5.jris.bean.Type
import de.bund.bfr.knime.fsklab.rakip.*
import ezvcard.VCard
import ezvcard.parameter.TelephoneType
import ezvcard.property.Address
import ezvcard.property.Email
import ezvcard.property.Organization
import ezvcard.property.StructuredName
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook


class ReadSheet {

    val workbook = XSSFWorkbook(
            this.javaClass.getResourceAsStream("simple_sheet.xlsx"))
}

enum class Column { A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z }

val RIS_TYPES = mapOf<String, Type>(
        "Abstract" to Type.ABST,
        "Audiovisual material" to Type.ADVS,
        "Aggregated Database" to Type.AGGR,
        "Ancient Text" to Type.ANCIENT,
        "Art Work" to Type.ART,
        "Bill" to Type.BILL,
        "Blog" to Type.BLOG,
        "Whole book" to Type.BOOK,
        "Case" to Type.CASE,
        "Book chapter" to Type.CHAP,
        "Chart" to Type.CHART,
        "Classical Work" to Type.CLSWK,
        "Computer program" to Type.COMP,
        "Conference proceeding" to Type.CONF,
        "Conference paper" to Type.CPAPER,
        "Catalog" to Type.CTLG,
        "Data file" to Type.DATA,
        "Online Database" to Type.DBASE,
        "Dictionary" to Type.DICT,
        "Electronic Book" to Type.EBOOK,
        "Electronic Book Section" to Type.ECHAP,
        "Edited Book" to Type.EDBOOK,
        "Electronic Article" to Type.EJOUR,
        "Web Page" to Type.ELEC,
        "Encyclopedia" to Type.ENCYC,
        "Equation" to Type.EQUA,
        "Figure" to Type.FIGURE,
        "Generic" to Type.GEN,
        "Government Document" to Type.GOVDOC,
        "Grant" to Type.GRANT,
        "Hearing" to Type.HEAR,
        "Internet Communication" to Type.ICOMM,
        "In Press" to Type.INPR,
        "Journal (full)" to Type.JFULL,
        "Journal" to Type.JOUR,
        "Legal Rule or Regulation" to Type.LEGAL,
        "Manuscript" to Type.MANSCPT,
        "Map" to Type.MAP,
        "Magazine article" to Type.MGZN,
        "Motion picture" to Type.MPCT,
        "Online Multimedia" to Type.MULTI,
        "Music score" to Type.MUSIC,
        "Newspaper" to Type.NEWS,
        "Pamphlet" to Type.PAMP,
        "Patent" to Type.PAT,
        "Personal communication" to Type.PCOMM,
        "Report" to Type.RPRT,
        "Serial publication" to Type.SER,
        "Slide" to Type.SLIDE,
        "Sound recording" to Type.SOUND,
        "Standard" to Type.STAND,
        "Statute" to Type.STAT,
        "Thesis/Dissertation" to Type.THES,
        "Unpublished work" to Type.UNPB,
        "Video recording" to Type.VIDEO
)

fun main(args: Array<String>) {
    val workbook = ReadSheet().workbook
    val sheet = workbook.getSheetAt(0)

//    val gm = sheet.retrieveGeneralInformation()
    val scope = sheet.importScope()
    print(scope)
}

/**
 * @throws IllegalStateException if the cell contains a string
 * @return 0 for blank cells
 */
fun XSSFSheet.getNumericValue(row: Int, col: Column): Double {
    val cell = getRow(row).getCell(col.ordinal)
    return cell.numericCellValue
}

/**
 * @return empty string for blank cells
 */
fun XSSFSheet.getStringValue(row: Int, col: Column): String {
    val cell = getRow(row).getCell(col.ordinal)
    return cell.stringCellValue
}

/**
 * Get strings from a cell with multiple values separated with commas.
 */
fun XSSFSheet.getStringListValue(row: Int, col: Column): List<String> {
    val cell = getRow(row).getCell(col.ordinal)
    return cell.stringCellValue.split(',')
}

/**
 * Import GeneralInformation from Excel sheet.
 *
 * - Model name cell in H2. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Source cell in H3. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Identifier cell in H4. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Creators in K4:Q7.
 * - Rights cell in H8. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Availability in H9. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 *   It only takes Yes and No.
 * - References in K14:U18.
 * - Language in H24. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Software in H25. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Language written in H26. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Model category in H27, H28, H29 and H32.
 * - Status in H33. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Objective in H34. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 * - Description in H35. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
 */
fun XSSFSheet.retrieveGeneralInformation(): GeneralInformation {

    /**
     * Import VCard from Excel row.
     *
     * - Name in the K column. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     *   Has the format: family name,given name
     * - Organization in the L column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     *   Only include organization. Unit names are not included.
     * - Telephone in the M column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     * - Mail in the N column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Mandatory.
     * - Country in the O column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional
     * - City in the P column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     * - Postal code in the Q column. [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC]. Optional.
     *
     * @throws IllegalStateException if mail cell is blank
     */
    fun XSSFSheet.importCreator(row: Int): VCard {

        val vCard = VCard()

        val nameText = getStringValue(row, Column.K)
        val organizationText = getStringValue(row, Column.L)
        val telephoneText = getStringValue(row, Column.M)
        val mailText = getStringValue(row, Column.N)
        val countryText = getStringValue(row, Column.O)
        val cityText = getStringValue(row, Column.P)
        val postalCodeInt = getNumericValue(row, Column.Q).toInt()

        // throw exception if mail is missing.
        if (mailText.isBlank())
            throw IllegalArgumentException("Missing mail")

        // name is optional. Ignore empty cell.
        if (nameText.isNotEmpty()) {
            val structuredName = StructuredName()
            structuredName.family = nameText.split(',')[0] // Assign family name
            structuredName.given = nameText.split(',')[1] // Assign given name
            vCard.structuredName = structuredName
        }

        // organization is optional. Ignore empty cell.
        if (organizationText.isNotEmpty()) {
            val organization = Organization()
            organization.values.add(organizationText)
            vCard.organization = organization
        }

        // telephone is optional. Ignore empty cell.
        if (telephoneText.isNotEmpty()) {
            vCard.addTelephoneNumber(telephoneText, TelephoneType.VOICE)
        }

        vCard.addEmail(Email(mailText))

        if (countryText.isNotEmpty() && cityText.isNotEmpty() && postalCodeInt != 0) {
            // import country, city and postal code
            val address = Address()
            address.country = countryText
            address.locality = cityText
            address.postalCode = postalCodeInt.toString()
            vCard.addAddress(address)
        }

        return vCard
    }

    /**
     * Import reference from Excel row.
     *
     * @throws IllegalArgumentException if isReferenceDescription or DOI are missing
     *
     * - Is_reference_description? in the K column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Mandatory. Takes "Yes" or "No".
     *   Other strings are discarded.
     *
     * - Publication type in the L column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional. Takes the full name
     *   of a RIS reference type.
     *
     * - Date in the M column. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
     *   Optional. Format `YYYY/MM/DD/other info` where the fields are optional.
     *   Examples: `2017/11/16/noon`, `2017/11/16`, `2017/11`, `2017`.
     *
     * - PubMed Id in the N column. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC].
     *   Optional. Unique unsigned integer. Example: 20069275
     *
     * - DOI in the O column. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
     *   Mandatory. Example: 10.1056/NEJM199710303371801.
     *
     * - Publication author list in the P column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional. The authors are defined
     *   with last name, name and joined with semicolons.
     *   Example: `Ungaretti-Haberbeck,L;Plaza-Rodríguez,C;Desvignes,V`
     *
     * - Publication title in the Q column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     *
     * - Publication abstract in the R column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     *
     * - Publication journal/vol/issue in the S column.
     *   Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING]. Optional.
     *
     * - Publication status.  // TODO: publication status
     *
     * - Publication website. Type [org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING].
     *   Optional. Invalid urls are discarded.
     */
    fun XSSFSheet.importReference(row: Int): Record {

        val isReferenceDescriptionText = getStringValue(row, Column.K)
        val typeText = getStringValue(row, Column.L)
        val dateText = getNumericValue(row, Column.M)
        val pmidText = getNumericValue(row, Column.N)
        val doiText = getStringValue(row, Column.O)
        val authorListText = getStringValue(row, Column.P)
        val titleText = getStringValue(row, Column.Q)
        val abstractText = getStringValue(row, Column.R)
        val journalVolIssueText = getStringValue(row, Column.S)
        val statusText = getStringValue(row, Column.T)
        val websiteText = getStringValue(row, Column.U)

        // throw exception if mail is isReferenceDescription or doi are missing
        if (isReferenceDescriptionText.isBlank())
            throw IllegalArgumentException("Missing Is reference description?")
        if (doiText.isBlank())
            throw IllegalArgumentException("Missing DOI")

        val record = Record()

        record.userDefinable1 = isReferenceDescriptionText // Save isReferenceDescription in U1 tag (user definable #1)

        record.type = RIS_TYPES[typeText] // Save type in TY tag

        // TODO: save date

        record.userDefinable2 = pmidText.toString() // Save PMID in U1 tag (user definable #2)

        record.doi = doiText // Save doi in DO tag

        authorListText.split(delimiters = ';').forEach { record.addAuthor(it) } // Save authors list

        record.title = titleText // Save title in TI tag

        record.abstr = abstractText  // Save abstract in AB tag

        // TODO: save journal

        record.userDefinable3 = statusText // Save status in U2 tag (user definable #3)

        record.url = websiteText // Save website in UR tag

        return record
    }

    /**
     * Import [ModelCategory] from Excel sheet.
     *
     * - Model class from H27. Mandatory.
     * - Model sub class from H28. Optional.
     * - Model class comment from H29. Optional.
     * - Basic process from H32. Optional.
     *
     * @throws IllegalArgumentException if modelClass is missing.
     */
    fun XSSFSheet.importModelCategory(): ModelCategory {

        if (getRow(26).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Missing model class")

        return ModelCategory(
                modelClass = getStringValue(26, Column.H),
                modelSubClass = getStringListValue(27, Column.H).toMutableList(),
                modelClassComment = getStringValue(28, Column.H),
                basicProcess = getStringListValue(31, Column.H).toMutableList())
    }

    val gm = GeneralInformation()

    gm.name = getStringValue(1, Column.H)
    gm.source = getStringValue(2, Column.H)
    gm.identifier = getStringValue(3, Column.H)

    // creators
    for (numRow in 3..7) {
        val vCard = try {
            importCreator(numRow)
        } catch (exception: Exception) {
            exception.printStackTrace()
            continue
        }
        gm.creators.add(vCard)
    }

    // creation date

    // modification dates

    gm.rights = getStringValue(7, Column.H)

    val availabilityString = getStringValue(8, Column.H)  // "Yes" or "No"
    gm.isAvailable = when (availabilityString) {
        "Yes" -> true
        "No" -> false
        else -> throw RuntimeException("Wrong value for availability: $availabilityString")
    }

    // references
    for (numRow in 13..17) {
        val record = try {
            importReference(numRow)
        } catch (exception: Exception) {
            exception.printStackTrace()
            continue
        }
        gm.reference.add(record)
    }

    gm.language = getStringValue(23, Column.H)
    gm.software = getStringValue(24, Column.H)
    gm.languageWrittenIn = getStringValue(25, Column.H)

    try {
      gm.modelCategory = importModelCategory()
    } catch (exception: Exception) {
        exception.printStackTrace()
    }

    gm.status = getStringValue(32, Column.H)
    gm.objective = getStringValue(33, Column.H)
    gm.status = getStringValue(34, Column.H)

    return gm
}

/**
 * Import Scope from Excel sheet.
 */
fun XSSFSheet.importScope(): Scope {

    /**
     * Import Hazard.
     *
     * - Hazard type in H48 cell. Mandatory.
     * - Hazard name in H49 cell. Mandatory.
     * - Hazard description in H50 cell. Optional.
     * - Hazard  unit in H51 cell. Mandatory.
     * - Hazard adverse effect in H52 cell. Optional.
     * - Source of contamination in H53 cell. Optional.
     * - Benchmark dose in H54 cell. Optional.
     * - Maximum residue limit in H55 cell. Optional.
     * - No observed adverse effect level in H56 cell. Optional.
     * - Lowest observed adverse effect level in H57 cell. Optional.
     * - Acceptable operator exposure level in H58 cell. Optional.
     * - Acute reference dose in H59 cell. Optional.
     * - Acceptable daily intake in H60 cell. Optional.
     * - Hazard ind/sum in H61 cell. Optional.
     */
    fun XSSFSheet.importHazard(): Hazard {

        if (getRow(47).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard type is missing")

        if (getRow(48).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard name is missing")

        if (getRow(50).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard unit is missing")

        val hazard = Hazard()
        hazard.hazardType = getStringValue(47, Column.H)
        hazard.hazardName = getStringValue(48, Column.H)
        hazard.hazardDescription = getStringValue(49, Column.H)
        hazard.hazardUnit = getStringValue(50, Column.H)
        hazard.adverseEffect= getStringValue(51, Column.H)
        // TODO: source of contamination
        hazard.benchmarkDose = getStringValue(53, Column.H)
        hazard.maximumResidueLimit = getStringValue(54, Column.H)
        hazard.noObservedAdverse = getStringValue(55, Column.H)
        hazard.lowestObservedAdverse = getStringValue(56, Column.H)
        hazard.acceptableOperator = getStringValue(57, Column.H)
        hazard.acuteReferenceDose = getStringValue(58, Column.H)
        hazard.acceptableDailyIntake = getStringValue(59, Column.H)
        hazard.hazardIndSum = getStringValue(60, Column.H)

        return hazard
    }

    /**
     * Import PopulationGroup.
     *
     * @throws IllegalArgumentException if population name is missing.
     *
     * - Population name in H62 cell. Mandatory.
     * - Target population in H63 cell. Cardinality +.
     * - Population span in H64 cell. Cardinality *.
     * - Population description in H65 cell. Cardinality *.
     * - Population age in H66 cell. Cardinality *.
     * - Population gender in H67 cell. Cardinality 1.
     * - BMI in H68 cell. Cardinality *.
     * - Special diet groups in H69 cell. Cardinality *.
     * - Pattern consumption in H70 cell. Cardinality *.
     * - Region in H71 cell. Cardinality *.
     * - Country in H72 cell. Cardinality *.
     * - Risk and population risk factor in H73 cell. Cardinality *.
     * - Season in H74 cell. Cardinality *.
     */
    fun XSSFSheet.importPopulationGroup(): PopulationGroup {

        if (getRow(61).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Missing population name")

        val populationName = getStringValue(61, Column.H)
        val targetPopulation = getStringValue(62, Column.H)
        val populationSpan = getStringListValue(63, Column.H).toMutableList()
        val populationDescription = getStringListValue(64, Column.H).toMutableList()
        val populationAge = getStringListValue(65, Column.H).toMutableList()
        val populationGender = getStringValue(66, Column.H)
        val bmi = getStringListValue(67, Column.H).toMutableList()
        val specialDietGroups = getStringListValue(68, Column.H).toMutableList()
        val patternConsumption = getStringListValue(69, Column.H).toMutableList()
        val region = getStringListValue(70, Column.H).toMutableList()
        val country = getStringListValue(71, Column.H).toMutableList()
        val riskAndPopulationFactors = getStringListValue(72, Column.H).toMutableList()
        val seasons = getStringListValue(73, Column.H).toMutableList()

        val group = PopulationGroup(populationName = populationName, targetPopulation = targetPopulation,
                populationSpan = populationSpan, populationDescription = populationDescription,
                populationAge = populationAge, populationGender = populationGender,
                bmi = bmi, specialDietGroups = specialDietGroups, patternConsumption = patternConsumption,
                region = region, country = country, populationRiskFactor = riskAndPopulationFactors,
                season = seasons)
        return group
    }

    val scope = Scope()
    scope.hazard = importHazard()
    scope.populationGroup = importPopulationGroup()
    // TODO: General comment
    // TODO: Temporal information
    // TODO: Spatial information

    return scope
}