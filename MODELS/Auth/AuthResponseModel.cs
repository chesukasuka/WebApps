namespace WebApps.MODELS.Auth
{
    public class AuthResponseModel
    {
        public string? token {  get; set; }
        public string? user_email {  get; set; }
        public string? user_nicename { get; set; }
        public string? user_display_name { get; set; }
    }

    public class AuthErrResponseModel
    {
        public string? code { get; set; }
        public string? message { get; set; }
        public AuthErrDataResponseModel? data { get; set; }
    }

    public class AuthErrDataResponseModel
    {
        public int? status { get; set; }
    }
}
