const { ApolloServer, gql } = require ('apollo-server');
const fetch = require('node-fetch');
const fileIcons = [
    { name: 'accdb' },
    { name: 'csv' },
    { name: 'docx' },
    { name: 'dotx' },
    { name: 'mpp' },
    { name: 'mpt' },
    { name: 'odp' },
    { name: 'ods' },
    { name: 'odt' },
    { name: 'one' },
    { name: 'onepkg' },
    { name: 'onetoc' },
    { name: 'potx' },
    { name: 'ppsx' },
    { name: 'pptx' },
    { name: 'pub' },
    { name: 'vsdx' },
    { name: 'vssx' },
    { name: 'vstx' },
    { name: 'xls' },
    { name: 'xlsx' },
    { name: 'xltx' },
    { name: 'xsn' }
];

const baseURL = `https://graph.microsoft.com/v1.0/me`

const typeDefs = gql`
    type Item {
      icon: String
      id: ID
      name: String
      modified: String
      modifiedBy: String
      modifiedByUserId: String
      size: String
      url: String
      path: String
      type: String
    }

    type User{
        imageUrl: String
        jobTitle: String
        displayName: String
    }

    type Query {
        items(url: String): [Item]
        user: User
    }
`;

const getSizeAsString = (sizeInBytes) => {
    let size = sizeInBytes/1024;
    if(size<1024){
      size = size.toFixed(2) + " KB";
      return size;
    }
    size = size/1024;
    if(size<1024){
      size = size.toFixed(2) + " MB";
      return size;
    }
    size = size/1024;
    size = size.toFixed(2) + " GB";
    return size;
};

const resolvers = {
    Query: {
        items: async (parent,args,context) => {
            console.log(args.url);
            const url = args.url ? args.url : `${baseURL}/drive/root/children`;
            return await fetch(url, {
                method: "GET",
                headers: {authorization: context.authorization}
            })
            .then(res => res.json())
            .then(json => {
                let items = [];
                json['value'].forEach((value)=>{
                    let item = {
                        icon: '',
                        id: value.id,
                        name: value.name,
                        modified: new Date(value.lastModifiedDateTime).toString().slice(0,25),
                        modifiedBy: value.lastModifiedBy.user.displayName,
                        modifiedByUserId: value.lastModifiedBy.user.id,
                        size: getSizeAsString(value.size),
                        url: value.webUrl,
                        path: value.parentReference.path,
                        type: 'file' in value ? "file" : "folder"
                    };
                    let name = value.name;
                    for (let i=0; i< fileIcons.length ; i++){
                        if(name.endsWith('.' + fileIcons[i].name)){
                            item.icon = `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${fileIcons[i].name}_16x1.svg`;
                        }
                    }
                    items.push(item);
                })
                return items;
            });
        },
        // user: async (parent,args,context) => {
        //     console.log(args.url);
        //     const url = args.url;
        //     return await fetch(url, {
        //         method: "GET",
        //         headers: {authorization: context.authorization}
        //     })
        //     .then( async (response)=>{
        //         let photourl = url + "/photo/$value";
        //         return await fetch(photourl, {
        //             method: "GET",
        //             headers: {authorization: context.authorization}
        //         })
        //         .then((res) => (res.blob()))
        //         .then((blob) => {
        //             let urlCreator = window.URL;
        //             let imageUrl = urlCreator.createObjectURL(blob);
        //             response.imageUrl = imageUrl;
        //             return response;
        //         });
        //     })
        // }
    }
}

const server = new ApolloServer({
    typeDefs,
    resolvers,
    context: ({req})=> ({
        authorization: req.headers.authorization
    })
});

server.listen().then(({url})=>{
    console.log(`Server ready at ${url}`);
});