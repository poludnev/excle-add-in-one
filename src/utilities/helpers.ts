export function getColumnLetter(columnNumber: number): string {
  if (columnNumber < 1) {
    throw new Error("Column number must be 1 or greater");
  }

  let temp: number;
  let letter: string = "";
  let num: number = columnNumber - 1; // Convert to 0-based for calculation

  while (num >= 0) {
    temp = num % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    num = Math.floor(num / 26) - 1;
  }

  return letter;
}

