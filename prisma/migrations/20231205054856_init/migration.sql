-- CreateTable
CREATE TABLE "Vessel" (
    "id" SERIAL NOT NULL,
    "name" TEXT NOT NULL,
    "HP" INTEGER NOT NULL,
    "MonthlyConsmCuM" INTEGER NOT NULL,
    "MonthlyMeRH" INTEGER NOT NULL,
    "MeConH" INTEGER NOT NULL,
    "MeConDCuM" INTEGER NOT NULL,
    "Dg1RH" INTEGER NOT NULL,
    "Dg2RH" INTEGER NOT NULL,
    "Dg3RH" INTEGER NOT NULL,
    "Dg4RH" INTEGER NOT NULL,
    "Dg5RH" INTEGER NOT NULL,
    "AuxRHTotal" INTEGER NOT NULL,
    "NumAuxRD" INTEGER NOT NULL,
    "AuxConH" INTEGER NOT NULL,
    "AuxConDCuM" INTEGER NOT NULL,
    "EstimatedVesselConAvgSailingConsumptionCuM" INTEGER NOT NULL,
    "RatioHpL" INTEGER NOT NULL,
    "TotalDist" INTEGER NOT NULL,
    "TotalRunningHourMePort" INTEGER NOT NULL,
    "TotalRunningHourMeSTBD" INTEGER NOT NULL,
    "EstimatedMajorOverHaulMePort" INTEGER NOT NULL,
    "EstimatedMajorOverHaulMeSTBD" INTEGER NOT NULL,

    CONSTRAINT "Vessel_pkey" PRIMARY KEY ("id")
);
