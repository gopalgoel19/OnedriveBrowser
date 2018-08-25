const { ApolloServer, gql } = require ('apollo-server');
const fetch = require('node-fetch')

const baseURL = `https://graph.microsoft.com/v1.0/me`

const typeDefs = gql`

    type item {
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

    type user{
        imageUrl: String
        jobTitle: String
        displayName: String
    }

    type Query {
        items: [item]
        # users: [user]
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
            const url = args.url ? args.url : `${baseURL}/drive/root/children`;
            return await fetch(url, {
                method: "GET",
                headers: {authorization: context.authorization}
            })
            .then(res => res.json())
            .then(json => {
                let items = [];
                json['value'].forEach((value)=>{
                    items.push({
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
                    });
                })
                console.log(items);
                return items;
            });
        }
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