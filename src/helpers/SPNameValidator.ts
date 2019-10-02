// import { Platform, ValidationType } from './enums/enums';
// import { ISPNameValidator } from './interfaces/ISPNameValidator';

export enum Platform {
  'SharePoint 2013 - 2016',
  'SharePoint Online',
}

export enum ValidationType {
  'File - Folder',
  'ListName',
}

interface ISPNameValidator {
  checkName(name: string, type: ValidationType): boolean;
}

export default class SPNameValidator implements ISPNameValidator {
  private platform: Platform;

  constructor(platform: Platform) {
    this.platform = platform;
  }

  public checkName(name: string, type: ValidationType): boolean {
    return this.ContainsIllegalCharOrWord(name, type);
  }

  private ContainsIllegalCharOrWord(
    value: string,
    type: ValidationType
  ): boolean {
    const len: number = value.length;
    let matches: boolean = len >= 1 && len <= 254;
    const charset: string[] = this.illegalCharList(type);
    const wordset: string[] = this.illegalWordList(type);

    if (matches) {
      for (let i: number = 0; i < len; i++) {
        matches = this.MatchASCII(charset, value[i]);
        if (!matches) {
          return false;
        }
      }
      if (matches) {
        const findWord = wordset.indexOf(value.toUpperCase());
        switch (findWord) {
          case -1:
            matches = true;
            break;
          default:
            matches = wordset[findWord].length !== value.length ? true : false;
            break;
        }
      }
    }
    return matches;
  }

  private MatchASCII(charset: string[], char: string): boolean {
    let match: boolean = true;
    if (match) {
      match = charset.indexOf(char) === -1;
      if (!match) {
        return false;
      }
    }
    return match;
  }

  private illegalCharList(type: ValidationType): string[] {
    let illegalCharacters: string[] = [];

    switch (type) {
      case ValidationType['File - Folder']:
        illegalCharacters =
          this.platform === Platform['SharePoint 2013 - 2016']
            ? [
                '~',
                '"',
                '#',
                '%',
                '&',
                '*',
                ':',
                '<',
                '>',
                '?',
                '/',
                '\\',
                '{',
                '|',
                '}',
                '.',
              ]
            : ['~', '"', '*', ':', '<', '>', '?', '/', '\\', '|'];
        break;
      case ValidationType.ListName:
        illegalCharacters =
          this.platform === Platform['SharePoint 2013 - 2016']
            ? [
                '~',
                '"',
                '#',
                '%',
                '&',
                '*',
                ':',
                '<',
                '>',
                '?',
                '/',
                '\\',
                '{',
                '|',
                '}',
                '.',
              ]
            : ['~', '"', '*', ':', '<', '>', '?', '/', '\\', '|'];
        break;
    }
    return illegalCharacters;
  }

  private illegalWordList(type: ValidationType): string[] {
    let illegalWords: string[] = [];

    switch (type) {
      case ValidationType['File - Folder']:
        illegalWords = [
          'AUX',
          'PRN',
          'NUL',
          'CON',
          'COM0',
          'COM1',
          'COM2',
          'COM3',
          'COM4',
          'COM5',
          'COM6',
          'COM7',
          'COM8',
          'COM9',
          'LPT0',
          'LPT1',
          'LPT2',
          'LPT3',
          'LPT4',
          'LPT5',
          'LPT6',
          'LPT7',
          'LPT8',
          'LPT9',
        ];
        break;
      case ValidationType.ListName:
        illegalWords = [
          'AUX',
          'PRN',
          'NUL',
          'CON',
          'COM0',
          'COM1',
          'COM2',
          'COM3',
          'COM4',
          'COM5',
          'COM6',
          'COM7',
          'COM8',
          'COM9',
          'LPT0',
          'LPT1',
          'LPT2',
          'LPT3',
          'LPT4',
          'LPT5',
          'LPT6',
          'LPT7',
          'LPT8',
          'LPT9',
        ];
        break;
    }
    return illegalWords;
  }
}
