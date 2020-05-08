import { HubConnection, HubConnectionBuilder, HttpTransportType, LogLevel, MessageType } from '@aspnet/signalr'

let signalRTokenEndpoint = "";
let cloudIncrementEndpoint = "";
let cloudAddEndpoint = "";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Adds two numbers.
 * @customfunction CLOUD_ADD
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function cloudAdd(first: number, second: number): Promise<number> {
  return onAzure(cloudAddEndpoint, { number1: first, number2: second});
}


/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Increments a value once a second.
 * @customfunction CLOUD_INCREMENT
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export async function cloudIncrement(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>) {
  let result = 0;
  const timer = setInterval(async () => {
    result += await onAzure(cloudIncrementEndpoint, { "ticker": incrementBy });
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

async function onAzure(url: string, data: object): Promise<number> {
  let fetchOptions: RequestInit = {
    method: 'post',
    mode: 'cors',
    cache: 'no-cache',
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(data)
  };
  var response = await fetch(url, fetchOptions);
  return response.text().then(text => Number.parseInt(text));
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Displays the current time once a second.
 * @customfunction CONNECT_TO_SIGNALR
 * @param invocation Custom function handler
 */
export async function initSignalR(channel: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  try {
    const res = await getSignalRInfo();
    if(typeof res === "string") {
      var response = JSON.parse(res);
      var options = {
        accessTokenFactory: () => response.accessToken
      }
      var connection: HubConnection = new HubConnectionBuilder().withUrl(response.url, options).build();
      connection.on(channel, (message:any) => {
        console.log(message);
        invocation.setResult("Message received: " + message);
      });
      invocation.onCanceled = async () => { await connection.stop(); console.log("disconnected") };
      await connection.start();
      console.log("connected");
    }  
  }
  catch (error) {
    console.error(error);
  }
}

async function getSignalRInfo() {
  try {
    const res = await fetch(signalRTokenEndpoint);
    return await res.text();
  }
  catch (error) {
    return console.log(error);
  }
}