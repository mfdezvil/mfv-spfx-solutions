declare interface IModalNewFormExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
  ModalFormTitle: string;
}

declare module 'ModalNewFormExtensionCommandSetStrings' {
  const strings: IModalNewFormExtensionCommandSetStrings;
  export = strings;
}
