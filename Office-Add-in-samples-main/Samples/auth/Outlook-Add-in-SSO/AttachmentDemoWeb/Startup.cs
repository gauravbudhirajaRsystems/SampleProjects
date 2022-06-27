﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(AttachmentDemoWeb.Startup))]

namespace AttachmentDemoWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}