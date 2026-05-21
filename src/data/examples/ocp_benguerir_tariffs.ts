import type { GridTariff } from "../../core/models/gridTariff.ts";
import { FULL_TARIFF_DH_PER_KWH, OFF_PEAK_TARIFF_DH_PER_KWH, PEAK_TARIFF_DH_PER_KWH } from "../../core/constants/ocpDefaults.ts";

export const ocpBenguerirTariff: GridTariff = {
  tariffName: "OCP Benguerir editable TOU tariff",
  peakTariffDhPerKWh: PEAK_TARIFF_DH_PER_KWH,
  fullTariffDhPerKWh: FULL_TARIFF_DH_PER_KWH,
  offPeakTariffDhPerKWh: OFF_PEAK_TARIFF_DH_PER_KWH,
  peakHours: [{ startHour: 17, endHour: 23 }],
  fullHours: [{ startHour: 7, endHour: 17 }],
  offPeakHours: [{ startHour: 22, endHour: 7 }],
  wheelingValueDhPerKWh: OFF_PEAK_TARIFF_DH_PER_KWH,
  currency: "DH",
};
