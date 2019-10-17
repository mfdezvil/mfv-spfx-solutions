import { IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';

export class NewsTypeOptions {
    public static CategoriaA: string = "Categoría A";
    public static CategoriaB: string = "Categoría B";
}

export class DateStringsOptions {
    public static DayPickerStrings: IDatePickerStrings = {
        months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
      
        shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      
        days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
      
        shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
      
        goToToday: 'Go to today',
        prevMonthAriaLabel: 'Go to previous month',
        nextMonthAriaLabel: 'Go to next month',
        prevYearAriaLabel: 'Go to previous year',
        nextYearAriaLabel: 'Go to next year',
        closeButtonAriaLabel: 'Close date picker',
      
        isRequiredErrorMessage: 'Start date is required.',
      
        invalidInputErrorMessage: 'Invalid date format.'
    };
}