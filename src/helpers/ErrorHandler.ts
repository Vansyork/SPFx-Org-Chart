import { ProcessHttpClientResponseException } from "@pnp/pnpjs";
export interface ErrorObjectFormat {
    statusText: string;
}
export default abstract class ErrorHandler {
    public static handleError(error: ProcessHttpClientResponseException): Promise<any>;
    public static handleError(error: string): Promise<any>;
    public static handleError(error: any): Promise<any>;
    public static handleError(error: ProcessHttpClientResponseException | string | any): Promise<any> {
        console.log(error);
        if (isProcessHttpClientResponseException(error)) {
            return Promise.reject(error);
        }
        else if (typeof error == 'string') {
            return Promise.reject(<ErrorObjectFormat>{ statusText: error });
        }
        else if (exceptionHasMessage(error)) {
            return Promise.reject(<ErrorObjectFormat>{ statusText: error.message });
        }
        else {
            return Promise.reject(<ErrorObjectFormat>{ statusText: "No error handler defined" });
        }

        function isProcessHttpClientResponseException(arg: any): arg is ProcessHttpClientResponseException {
            if (arg.name === undefined) {
                return false;
            } else {
                return arg.name === "ProcessHttpClientResponseException";
            }
        }
        function exceptionHasMessage(arg: any) {
            if (arg.message === undefined) {
                return false;
            } else {
                return true;
            }
        }
    }
}