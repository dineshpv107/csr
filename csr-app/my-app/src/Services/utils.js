
class Util {

    getUrlParameterId() {
        return window.location.pathname?.substr(window.location.pathname?.lastIndexOf('/') + 1) || undefined; // .match(/\/([^/]*)$/)[1]
    };
}

export default Util;