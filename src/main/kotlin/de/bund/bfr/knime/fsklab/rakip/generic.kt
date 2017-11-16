/***************************************************************************************************
 * Copyright (c) 2015 Federal Institute for Risk Assessment (BfR), Germany
 * <p>
 * This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
 * License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
 * version.
 * <p>
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 * <p>
 * You should have received a copy of the GNU General Public License along with this program. If not, see
 * <http://www.gnu.org/licenses/>.
 * <p>
 * Contributors: Department Biological Safety - BfR
 **************************************************************************************************/
package de.bund.bfr.knime.fsklab.rakip

import com.gmail.gcolaianni5.jris.bean.Record
import ezvcard.VCard
import java.net.URI
import java.net.URL
import java.util.*
import ezvcard.Ezvcard
import com.fasterxml.jackson.core.JsonProcessingException
import java.io.IOException
import com.fasterxml.jackson.core.JsonParser
import com.fasterxml.jackson.core.JsonGenerator
import com.fasterxml.jackson.core.Version
import com.gmail.gcolaianni5.jris.exception.JRisException
import com.gmail.gcolaianni5.jris.engine.JRis
import java.io.StringReader
import java.util.Arrays
import java.io.StringWriter
import com.fasterxml.jackson.databind.*
import com.fasterxml.jackson.databind.module.SimpleModule


data class Assay(var name: String = "", var description: String = "")

data class DataBackground(
        var study: Study = Study(),
        var studySample: StudySample = StudySample(),
        var dietaryAssessmentMethod: DietaryAssessmentMethod = DietaryAssessmentMethod(),
        val laboratoryAccreditation: MutableList<String> = mutableListOf(),
        var assay: Assay = Assay())

data class DietaryAssessmentMethod(
        var collectionTool: String = "",
        var numberOfNonConsecutiveOneDay: Int = 0,
        var softwareTool: String = "",
        val numberOfFoodItems: MutableList<String> = mutableListOf(),
        val recordTypes: MutableList<String> = mutableListOf(),
        val foodDescriptors: MutableList<String> = mutableListOf())

/**
 * @property treatment Methodological treatment of left-censored data
 * @property contaminationLevel Level of contamination after left-censored data treatment
 * @property exposureType Type of exposure
 * @property scenario Scenario of exposure assessment
 * @property uncertaintyEstimation Analysis to estimate uncertainty
 */
data class Exposure(
        var treatment: String = "",
        var contaminationLevel: String = "",
        var exposureType: String = "",
        var scenario: String = "",
        var uncertaintyEstimation: String = "")

/**
 * @property name Name given to the model or data
 * @property source Related resource from which the resource is derived
 * @property identifier Unambiguous ID given to the model or data
 * @property creationDate Model creation date
 * @property rights Rights held in over the resource
 * @property isAvailable Availability of data or model
 * @property url Web address referencing the resource location
 * @property format Form of data (file extension)
 * @property language Language of the resource
 * @property software Program in which the model has been implemented
 * @property languageWrittenIn  Language used to write the model
 * @property status Curation status of the model
 * @property objective Objective of the model or data
 * @property description General assayDescription of the study, data or model
 */
data class GeneralInformation(
        var name: String = "",
        var source: String = "",
        var identifier: String = "",
        var creators: MutableList<VCard> = mutableListOf(),
        var creationDate: Date = Date(),
        var modificationDate: MutableList<Date> = mutableListOf(),
        var rights: String = "",
        var isAvailable: Boolean = false,
        var url: URL? = null,
        var format: String = "",
        var reference: MutableList<Record> = mutableListOf(),
        var language: String = "",
        var software: String = "",
        var languageWrittenIn: String = "",
        var modelCategory: ModelCategory = ModelCategory(),
        var status: String = "",
        var objective: String = "",
        var description: String = ""
)

data class GenericModel(
        var generalInformation: GeneralInformation = GeneralInformation(),
        var scope: Scope = Scope(),
        var dataBackground: DataBackground = DataBackground(),
        var modelMath: ModelMath = ModelMath(),
        var simulation: Simulation = Simulation()
)

data class Hazard(
        var hazardType: String = "",
        var hazardName: String = "",
        var hazardDescription: String = "",
        var hazardUnit: String = "",
        var adverseEffect: String = "",
        var origin: String = "",
        var benchmarkDose: String = "",
        var maximumResidueLimit: String = "",
        var noObservedAdverse: String = "",
        var lowestObservedAdverse: String = "",
        var acceptableOperator: String = "",
        var acuteReferenceDose: String = "",
        var acceptableDailyIntake: String = "",
        var hazardIndSum: String = "",
        var laboratoryName: String = "",
        var laboratoryCountry: String = "",
        var detectionLimit: String = "",
        var quantificationLimit: String = "",
        var leftCensoredData: String = "",
        var rangeOfContamination: String = "")

/**
 * @property modelClass Type of model. Ultimate goal of the global model.
 * @property modelSubClass Sub-classification of the model given modelClass.
 * @property basicProcess Impact of the specific process in the hazard.
 */
data class ModelCategory(
        var modelClass: String = "",
        val modelSubClass: MutableList<String> = mutableListOf(),
        var modelClassComment: String = "",
        val modelSubSubClass: MutableList<String> = mutableListOf(),
        val basicProcess: MutableList<String> = mutableListOf())

/**
 * @property equationName Model equation name
 * @property equationClass Model equation class
 * @property equationReference Model equation references (RIS)
 * @property equation Model equation or script
 */
data class ModelEquation(
        var equationName: String = "",
        var equationClass: String = "",
        val equationReference: MutableList<Record> = mutableListOf(),
        var equation: String = ""
)

data class ModelMath(
        val parameter: MutableList<Parameter> = mutableListOf(),
        var sse: Double = .0,
        var mse: Double = .0,
        var rmse: Double = .0,
        var rSquared: Double = .0,
        var aic: Double = .0,
        var bic: Double = .0,
        var modelEquation: MutableList<ModelEquation> = mutableListOf(),
        var fittingProcedure: String = "",
        var exposure: Exposure = Exposure(),
        val event: MutableList<String> = mutableListOf()
)

class Parameter(
        var id: String = "",
        var classification: Classification = Classification.constant,
        var name: String = "",
        var description: String = "",
        var unit: String = "",
        var unitCategory: String = "",
        var dataType: String = "",
        var source: String = "",
        var subject: String = "",
        var distribution: String = "",
        var value: String = "",
        var reference: String = "",
        var variabilitySubject: String = "",
        val modelApplicability: MutableList<String> = mutableListOf(),
        var error: Double = .0
) {
    enum class Classification { input, output, constant }
}

data class PopulationGroup(
        var populationName: String = "",
        var targetPopulation: String = "",
        val populationSpan: MutableList<String> = mutableListOf(),
        val populationDescription: MutableList<String> = mutableListOf(),
        val populationAge: MutableList<String> = mutableListOf(),
        var populationGender: String = "",
        val bmi: MutableList<String> = mutableListOf(),
        val specialDietGroups: MutableList<String> = mutableListOf(),
        val patternConsumption: MutableList<String> = mutableListOf(),
        val region: MutableList<String> = mutableListOf(),
        val country: MutableList<String> = mutableListOf(),
        val populationRiskFactor: MutableList<String> = mutableListOf(),
        val season: MutableList<String> = mutableListOf()
)

data class Product(
        var environmentName: String = "",
        var environmentDescription: String = "",
        var environmentUnit: String = "",
        val productionMethod: MutableList<String> = mutableListOf(),
        val packaging: MutableList<String> = mutableListOf(),
        val productTreatment: MutableList<String> = mutableListOf(),
        var originCountry: String = "",
        var areaOfOrigin: String = "",
        var fisheriesArea: String = "",
        var productionDate: Date = Date(),
        var expirationDate: Date = Date()
)

// TODO("RakipModule")

/**
 * @property generalComment General comment on the model
 * @property temporalInformation Temporal information on which the model applies
 * @property region information on which the model applies
 * @property country countries on which the model applies
 */
data class Scope(
        var product: Product = Product(),
        var hazard: Hazard = Hazard(),
        var populationGroup: PopulationGroup = PopulationGroup(),
        var generalComment: String = "",
        var temporalInformation: String = "",
        val region: MutableList<String> = mutableListOf(),
        val country: MutableList<String> = mutableListOf()
)

/**
 * @property algorithm Simulation algorithm
 * @property simulatedModel URI of the model applied for simulation/prediction
 * @property parameterSettings
 * @property description General assayDescription of the simulation
 * @property plotSettings Information on the parameters to be plotted
 * @property visualizationScript Pointer to software code (R script)
 */
data class Simulation(
        var algorithm: String = "",
        var simulatedModel: String = "",
        val parameterSettings: MutableList<String> = mutableListOf(),
        var description: String = "",
        val plotSettings: MutableList<String> = mutableListOf(),
        var visualizationScript: String = ""
)


/**
 * @property title Study title
 * @property description Study assayDescription
 * @property designType Study type
 * @property measurementType Observed measure in the assay
 * @property technologyType Employed technology to observe this measurement
 * @property technologyPlatform Used technology platform
 * @property accreditationProcedure Used accreditation procedure
 * @property protocolDescription Type of the protocol (e.g. Extraction protocol)
 * @property parametersName Parameters used when executing this protocol
 */
data class Study(
        var title: String = "",
        var description: String = "",
        var designType: String = "",
        var measurementType: String = "",
        var technologyType: String = "",
        var technologyPlatform: String = "",
        var accreditationProcedure: String = "",
        var protocolName: String = "",
        var protocolType: String = "",
        var protocolDescription: String = "",
        var protocolURI: URI? = null,
        var protocolVersion: String = "",
        var parametersName: String = "",
        var componentsName: String = "",
        var componentsType: String = "")

data class StudySample(
        var sample: String = "",
        var moisturePercentage: Double = .0,
        var fatPercentage: Double = .0,
        var collectionProtocol: String = "",
        var samplingStrategy: String = "",
        var samplingProgramType: String = "",
        var samplingMethod: String = "",
        var samplingPlan: String = "",
        var samplingWeight: String = "",
        var samplingSize: String = "",
        var lotSizeUnit: String = "",
        var samplingPoint: String = ""
)

class RakipModule : SimpleModule("RakipModule", Version.unknownVersion()) {
    init {

        addSerializer(Record::class.java, object : JsonSerializer<Record>() {
            @Throws(IOException::class, JsonProcessingException::class)
            override fun serialize(value: Record, gen: JsonGenerator, serializers: SerializerProvider) {

                gen.writeStartObject()

                try {
                    StringWriter().use { writer ->
                        JRis.build(Arrays.asList(value), writer)
                        gen.writeStringField("reference", writer.toString())
                    }
                } catch (e: JRisException) {
                    throw IOException(e)
                }

                gen.writeEndObject()
            }
        })

        addDeserializer(Record::class.java, object : JsonDeserializer<Record>() {
            @Throws(IOException::class, JsonProcessingException::class)
            override fun deserialize(parser: JsonParser, context: DeserializationContext): Record {

                val node = parser.readValueAsTree<JsonNode>()
                val referenceString = node.get("reference").asText()

                try {
                    StringReader(referenceString).use { reader -> return JRis.parse(reader)[0] }
                } catch (e: JRisException) {
                    throw IOException(e)
                }
            }
        })

        addSerializer(VCard::class.java, object : JsonSerializer<VCard>() {
            @Throws(IOException::class, JsonProcessingException::class)
            override fun serialize(value: VCard, gen: JsonGenerator, provider: SerializerProvider) {

                gen.writeStartObject()
                gen.writeStringField("creator", value.writeJson())
                gen.writeEndObject()
            }
        })

        addDeserializer(VCard::class.java, object : JsonDeserializer<VCard>() {
            @Throws(IOException::class, JsonProcessingException::class)
            override fun deserialize(parser: JsonParser, context: DeserializationContext): VCard {

                val node = parser.readValueAsTree<JsonNode>()
                val cardString = node.get("creator").asText()
                return Ezvcard.parseJson(cardString).first()
            }
        })
    }
}