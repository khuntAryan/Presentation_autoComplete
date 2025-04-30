"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.logDebug = exports.log = exports.Logger = exports.last = exports.vd = exports.GeneralHelper = void 0;
class GeneralHelper {
    static arrayify(input) {
        if (Array.isArray(input)) {
            return input;
        }
        else if (input !== undefined) {
            return [input];
        }
        else {
            return [];
        }
    }
    static propertyExists(object, property) {
        if (!object || typeof object !== 'object')
            return false;
        return !!Object.getOwnPropertyDescriptor(object, property);
    }
}
exports.GeneralHelper = GeneralHelper;
const vd = (v, keys) => {
    if (keys && typeof v === 'object') {
        v = Object.keys(v);
    }
    console.log('--------- [pptx-automizer] ---------');
    // @ts-ignore
    console.log(new Error().stack.split('\n')[2].trim());
    console.dir(v, { depth: 10 });
};
exports.vd = vd;
const last = (arr) => arr[arr.length - 1];
exports.last = last;
exports.Logger = {
    verbosity: 1,
    target: 'console',
    log: (message, verbosity, showStack, target) => {
        if (verbosity > exports.Logger.verbosity) {
            return;
        }
        target = target || exports.Logger.target;
        if (target === 'console') {
            if (showStack) {
                (0, exports.vd)(message);
            }
            else {
                console.log(message);
            }
        }
        else {
            // TODO: append message to a logfile
        }
    },
};
const log = (message, verbosity) => {
    exports.Logger.log(message, verbosity);
};
exports.log = log;
const logDebug = (message, verbosity) => {
    exports.Logger.log(message, verbosity, true);
};
exports.logDebug = logDebug;
//# sourceMappingURL=general-helper.js.map