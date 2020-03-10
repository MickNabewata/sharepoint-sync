import * as moment from "moment";

/** Excelのシリアル値を日付型に変換 */
export function toDate(value: number): Date {
  let m = moment((value - (25567 + 2)) * 86400 * 1000);
  const d = moment.duration(m.utcOffset(), "minutes");
  m = m.subtract(d);
  return m && m.isValid() ? m.toDate() : undefined;
}