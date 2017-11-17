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

enum class Row(val num: Int) {

    GENERALINFORMATION_NAME(1),
    GENERALINFORMATION_SOURCE(2),
    GENERALINFORMATION_IDENTIFIER(3),
    GENERALINFORMATION_RIGHTS(7),
    GENERALINFORMATION_AVAILABILITY(8),
    GENERALINFORMATION_LANGUAGE(23),
    GENERALINFORMATION_SOFTWARE(24),
    GENERALINFORMATION_LANGUAGEWRITTENIN(25),
    GENERALINFORMATION_STATUS(33),
    GENERALINFORMATION_OBJECTIVE(34),
    GENERALINFORMATION_DESCRIPTION(35),

    MODELCATEGORY_MODELCLASS(26),
    MODELCATEGORY_MODELSUBCLASS(27),
    MODELCATEGORY_MODELCLASSCOMMENT(28),
    MODELCATEGORY_BASICPROCESS(31),

    HAZARD_TYPE(47),
    HAZARD_NAME(48),
    HAZARD_DESCRIPTION(49),
    HAZARD_UNIT(50),
    HAZARD_ADVERSEEFFECT(51),
    HAZARD_BENCHMARKDOSE(53),
    HAZARD_MAXIMUMRESIDUELIMIT(54),
    HAZARD_NOOBSERVEDADVERSE(55),
    HAZARD_LOWESTOBSERVEDADVERSE(56),
    HAZARD_ACCEPTABLEOPERATOR(57),
    HAZARD_ACUTEREFERENCEDOSE(58),
    HAZARD_ACCEPTABLEDAILYINTAKE(59),
    HAZARD_INDSUM(60),

    POPULATIONGROUP_NAME(61),
    POPULATIONGROUP_TARGET(62),
    POPULATIONGROUP_SPAN(63),
    POPULATIONGROUP_DESCRIPTION(64),
    POPULATIONGROUP_AGE(65),
    POPULATIONGROUP_GENDER(66),
    POPULATIONGROUP_BMI(67),
    POPULATIONGROUP_SPECIALDIETGROUPS(68),
    POPULATIONGROUP_PATTERNCONSUMPTION(69),
    POPULATIONGROUP_REGION(70),
    POPULATIONGROUP_COUNTRY(71),
    POPULATIONGROUP_RISKFACTOR(72),
    POPULATIONGROUP_SEASON(73),

    SCOPE_COMMENT(74)
}

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
 * @throws IllegalStateException if the cell contains a string
 * @return 0 for blank cells
 */
fun XSSFSheet.getNumericValue(row: Row, col: Column): Double = getNumericValue(row.num, col)

/**
 * @return empty string for blank cells
 */
fun XSSFSheet.getStringValue(row: Int, col: Column): String {
    val cell = getRow(row).getCell(col.ordinal)
    return cell.stringCellValue
}

/**
 * @return empty string for blank cells
 */
fun XSSFSheet.getStringValue(row: Row, col: Column): String = getStringValue(row.num, col)

/**
 * Get strings from a cell with multiple values separated with commas.
 */
fun XSSFSheet.getStringListValue(row: Int, col: Column): List<String> {
    val cell = getRow(row).getCell(col.ordinal)
    return cell.stringCellValue.split(',')
}

/**
 * Get strings from a cell with multiple values separated with commas.
 */
fun XSSFSheet.getStringListValue(row: Row, col: Column): List<String> = getStringListValue(row.num, col)

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

        if (getRow(Row.MODELCATEGORY_MODELCLASS.num).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Missing model class")

        return ModelCategory(
                modelClass = getStringValue(Row.MODELCATEGORY_MODELCLASS, Column.H),
                modelSubClass = getStringListValue(Row.MODELCATEGORY_MODELCLASS, Column.H).toMutableList(),
                modelClassComment = getStringValue(Row.MODELCATEGORY_MODELCLASSCOMMENT, Column.H),
                basicProcess = getStringListValue(Row.MODELCATEGORY_BASICPROCESS, Column.H).toMutableList())
    }

    val gm = GeneralInformation()

    gm.name = getStringValue(Row.GENERALINFORMATION_NAME, Column.H)
    gm.source = getStringValue(Row.GENERALINFORMATION_SOURCE, Column.H)
    gm.identifier = getStringValue(Row.GENERALINFORMATION_IDENTIFIER, Column.H)

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

    gm.rights = getStringValue(Row.GENERALINFORMATION_RIGHTS, Column.H)

    val availabilityString = getStringValue(Row.GENERALINFORMATION_AVAILABILITY, Column.H)  // "Yes" or "No"
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

    gm.language = getStringValue(Row.GENERALINFORMATION_LANGUAGE, Column.H)
    gm.software = getStringValue(Row.GENERALINFORMATION_SOFTWARE, Column.H)
    gm.languageWrittenIn = getStringValue(Row.GENERALINFORMATION_LANGUAGEWRITTENIN, Column.H)

    try {
        gm.modelCategory = importModelCategory()
    } catch (exception: Exception) {
        exception.printStackTrace()
    }

    gm.status = getStringValue(Row.GENERALINFORMATION_STATUS, Column.H)
    gm.objective = getStringValue(Row.GENERALINFORMATION_OBJECTIVE, Column.H)
    gm.description = getStringValue(Row.GENERALINFORMATION_DESCRIPTION, Column.H)

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

        if (getRow(Row.HAZARD_TYPE.num).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard type is missing")

        if (getRow(Row.HAZARD_NAME.num).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard name is missing")

        if (getRow(Row.HAZARD_UNIT.num).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Hazard unit is missing")

        return Hazard(
                hazardType = getStringValue(Row.HAZARD_TYPE, Column.H),
                hazardName = getStringValue(Row.HAZARD_NAME, Column.H),
                hazardDescription = getStringValue(Row.HAZARD_DESCRIPTION, Column.H),
                hazardUnit = getStringValue(Row.HAZARD_UNIT, Column.H),
                adverseEffect = getStringValue(Row.HAZARD_ADVERSEEFFECT, Column.H),
                // TODO: source of contamination
                benchmarkDose = getStringValue(Row.HAZARD_BENCHMARKDOSE, Column.H),
                maximumResidueLimit = getStringValue(Row.HAZARD_MAXIMUMRESIDUELIMIT, Column.H),
                noObservedAdverse = getStringValue(Row.HAZARD_NOOBSERVEDADVERSE, Column.H),
                lowestObservedAdverse = getStringValue(Row.HAZARD_LOWESTOBSERVEDADVERSE, Column.H),
                acceptableOperator = getStringValue(Row.HAZARD_ACCEPTABLEOPERATOR, Column.H),
                acuteReferenceDose = getStringValue(Row.HAZARD_ACUTEREFERENCEDOSE, Column.H),
                acceptableDailyIntake = getStringValue(Row.HAZARD_ACCEPTABLEDAILYINTAKE, Column.H),
                hazardIndSum = getStringValue(Row.HAZARD_INDSUM, Column.H))
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

        if (getRow(Row.POPULATIONGROUP_NAME.num).getCell(Column.H.ordinal).cellType == Cell.CELL_TYPE_BLANK)
            throw IllegalArgumentException("Missing population name")

        return PopulationGroup(
                populationName = getStringValue(Row.POPULATIONGROUP_NAME, Column.H),
                targetPopulation = getStringValue(Row.POPULATIONGROUP_TARGET, Column.H),
                populationSpan = getStringListValue(Row.POPULATIONGROUP_SPAN, Column.H).toMutableList(),
                populationDescription = getStringListValue(Row.POPULATIONGROUP_DESCRIPTION, Column.H).toMutableList(),
                populationAge = getStringListValue(Row.POPULATIONGROUP_AGE, Column.H).toMutableList(),
                populationGender = getStringValue(Row.POPULATIONGROUP_GENDER, Column.H),
                bmi = getStringListValue(Row.POPULATIONGROUP_BMI, Column.H).toMutableList(),
                specialDietGroups = getStringListValue(Row.POPULATIONGROUP_SPECIALDIETGROUPS, Column.H).toMutableList(),
                patternConsumption = getStringListValue(Row.POPULATIONGROUP_SPECIALDIETGROUPS, Column.H).toMutableList(),
                region = getStringListValue(Row.POPULATIONGROUP_REGION, Column.H).toMutableList(),
                country = getStringListValue(Row.POPULATIONGROUP_COUNTRY, Column.H).toMutableList(),
                populationRiskFactor = getStringListValue(Row.POPULATIONGROUP_RISKFACTOR, Column.H).toMutableList(),
                season = getStringListValue(Row.POPULATIONGROUP_SEASON, Column.H).toMutableList())
    }

    val scope = Scope()
    scope.hazard = importHazard()
    scope.populationGroup = importPopulationGroup()
    scope.generalComment = getStringValue(Row.SCOPE_COMMENT, Column.H)

    // TODO: Temporal information
    // TODO: Spatial information

    return scope
}
