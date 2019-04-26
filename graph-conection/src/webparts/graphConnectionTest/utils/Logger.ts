import {Constants } from './Constants';

export class Logger {

    public static WriteError(functionName: string, message: string){
        console.error(Constants.AppName+" - ERROR at "+functionName+": "+message);
    } 
}